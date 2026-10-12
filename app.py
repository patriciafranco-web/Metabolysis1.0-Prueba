from flask import Flask, render_template, request, redirect, url_for, Response, send_file, session
from flask_sqlalchemy import SQLAlchemy
from werkzeug.security import generate_password_hash, check_password_hash
from functools import wraps
import io
import csv
import openpyxl
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas

app = Flask(__name__)
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///metabolisys_pro.db'
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
app.secret_key = 'metabolisys_secreto_seguro_2026'
db = SQLAlchemy(app)

class User(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(50), unique=True, nullable=False)
    password_hash = db.Column(db.String(200), nullable=False)

class Patient(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    full_name = db.Column(db.String(100), nullable=False)
    ci = db.Column(db.String(20), unique=True, nullable=False)
    age = db.Column(db.Integer, nullable=False)
    gender = db.Column(db.String(10), nullable=False)
    phone = db.Column(db.String(20))
    sanatorio = db.Column(db.String(100), nullable=False)
    pathology_type = db.Column(db.String(100))
    evolution_years = db.Column(db.Integer)
    comorbidities = db.Column(db.String(200))
    records = db.relationship('MedicalRecord', backref='patient', lazy=True, cascade="all, delete-orphan")

class MedicalRecord(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    patient_id = db.Column(db.Integer, db.ForeignKey('patient.id'), nullable=False)
    date = db.Column(db.String(20), nullable=False)
    weight = db.Column(db.Float)
    height = db.Column(db.Float)
    waist = db.Column(db.Float)
    body_fat = db.Column(db.Float)
    muscle_mass = db.Column(db.Float)
    visceral_fat = db.Column(db.Float)
    glucose = db.Column(db.Float)
    insulin = db.Column(db.Float)
    hba1c = db.Column(db.Float)
    cholesterol_tot = db.Column(db.Float)
    hdl = db.Column(db.Float)
    ldl = db.Column(db.Float)
    triglycerides = db.Column(db.Float)
    bp_systolic = db.Column(db.Integer)
    bp_diastolic = db.Column(db.Integer)
    creatinine = db.Column(db.Float)
    egfr = db.Column(db.Float)
    albuminuria = db.Column(db.Float)
    alt_tgp = db.Column(db.Float)
    ast_tgo = db.Column(db.Float)
    medications = db.Column(db.String(300))
    motivo_consulta = db.Column(db.Text)
    examen_fisico = db.Column(db.Text)
    notes = db.Column(db.Text)
    conducta_medica = db.Column(db.Text)

    @property
    def bmi(self):
        if self.weight and self.height and self.height > 0:
            return round(self.weight / (self.height ** 2), 2)
        return 0

    @property
    def homair(self):
        if self.glucose and self.insulin:
            return round((self.glucose * self.insulin) / 405, 2)
        return None

    @property
    def waist_height_ratio(self):
        if self.waist and self.height and self.height > 0:
            height_cm = self.height * 100
            return round(self.waist / height_cm, 2)
        return None

    @property
    def basal_metabolic_rate(self):
        if self.weight and self.height and self.patient and self.patient.age:
            height_cm = self.height * 100
            weight_kg = self.weight
            age = self.patient.age
            if self.patient.gender == 'Hombre':
                return round((10 * weight_kg) + (6.25 * height_cm) - (5 * age) + 5, 1)
            else:
                return round((10 * weight_kg) + (6.25 * height_cm) - (5 * age) - 161, 1)
        return None

with app.app_context():
    db.create_all()
    # Crear usuario de prueba por defecto si no existe
    if not User.query.filter_by(username='dra.metabolisys').first():
        hashed_pw = generate_password_hash('metabo2026')
        admin = User(username='dra.metabolisys', password_hash=hashed_pw)
        db.session.add(admin)
        db.session.commit()

def login_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if 'user_id' not in session:
            return redirect(url_for('login'))
        return f(*args, **kwargs)
    return decorated_function

@app.route('/login', methods=['GET', 'POST'])
def login():
    error = None
    if request.method == 'POST':
        username = request.form.get('username', '').strip()
        password = request.form.get('password', '').strip()
        
        # Acceso directo y ultra simple
        if username == 'admin' and password == '123':
            session['user_id'] = 1
            session['username'] = 'admin'
            return redirect(url_for('index'))
        else:
            error = 'Usuario o contraseña incorrectos. Usa: admin / 123'
            
    return render_template('login.html', error=error)



@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login'))

@app.route('/')
@login_required
def index():
    search_query = request.args.get('q', '').strip()
    sanatorio_filter = request.args.get('sanatorio', '').strip()
    
    query = Patient.query
    if sanatorio_filter:
        query = query.filter_by(sanatorio=sanatorio_filter)
    if search_query:
        query = query.filter(
            (Patient.full_name.ilike(f'%{search_query}%')) | 
            (Patient.ci.ilike(f'%{search_query}%'))
        )
        
    patients = query.all()
    sanatorios = db.session.query(Patient.sanatorio).distinct().all()
    sanatorios = [s[0] for s in sanatorios if s[0]]

    return render_template('index.html', patients=patients, sanatorios=sanatorios, current_sanatorio=sanatorio_filter, search_query=search_query)

@app.route('/patient/new', methods=['GET', 'POST'])
@login_required
def new_patient():
    if request.method == 'POST':
        patient = Patient(
            full_name=request.form['full_name'],
            ci=request.form['ci'],
            age=int(request.form['age']),
            gender=request.form['gender'],
            phone=request.form['phone'],
            sanatorio=request.form['sanatorio'],
            pathology_type=request.form['pathology_type'],
            evolution_years=int(request.form['evolution_years'] or 0),
            comorbidities=request.form['comorbidities']
        )
        db.session.add(patient)
        db.session.commit()
        return redirect(url_for('index'))
    return render_template('new_patient.html')

@app.route('/patient/<int:id>')
@login_required
def patient_detail(id):
    patient = Patient.query.get_or_404(id)
    dates = [r.date for r in patient.records]
    hba1c_data = [r.hba1c for r in patient.records]
    weight_data = [r.weight for r in patient.records]

    return render_template(
        'patient_detail.html', 
        patient=patient, 
        dates=dates, 
        hba1c_data=hba1c_data, 
        weight_data=weight_data
    )

@app.route('/record/new/<int:patient_id>', methods=['GET', 'POST'])
@login_required
def new_record(patient_id):
    patient = Patient.query.get_or_404(patient_id)
    if request.method == 'POST':
        record = MedicalRecord(
            patient_id=patient.id,
            date=request.form['date'],
            weight=float(request.form['weight'] or 0),
            height=float(request.form['height'] or 0),
            waist=float(request.form['waist'] or 0),
            body_fat=float(request.form['body_fat'] or 0),
            muscle_mass=float(request.form['muscle_mass'] or 0),
            visceral_fat=float(request.form['visceral_fat'] or 0),
            glucose=float(request.form['glucose'] or 0),
            insulin=float(request.form['insulin'] or 0),
            hba1c=float(request.form['hba1c'] or 0),
            cholesterol_tot=float(request.form['cholesterol_tot'] or 0),
            hdl=float(request.form['hdl'] or 0),
            ldl=float(request.form['ldl'] or 0),
            triglycerides=float(request.form['triglycerides'] or 0),
            bp_systolic=int(request.form['bp_systolic'] or 0),
            bp_diastolic=int(request.form['bp_diastolic'] or 0),
            creatinine=float(request.form['creatinine'] or 0),
            egfr=float(request.form['egfr'] or 0),
            albuminuria=float(request.form['albuminuria'] or 0),
            alt_tgp=float(request.form['alt_tgp'] or 0),
            ast_tgo=float(request.form['ast_tgo'] or 0),
            medications=request.form['medications'],
            motivo_consulta=request.form['motivo_consulta'],
            examen_fisico=request.form['examen_fisico'],
            notes=request.form['notes'],
            conducta_medica=request.form['conducta_medica']
        )
        db.session.add(record)
        db.session.commit()
        return redirect(url_for('patient_detail', id=patient.id))
    return render_template('new_record.html', patient=patient)

@app.route('/import/excel', methods=['POST'])
@login_required
def import_excel():
    if 'excel_file' not in request.files:
        return redirect(url_for('index'))
    file = request.files['excel_file']
    if file.filename == '':
        return redirect(url_for('index'))
    
    if file and file.filename.endswith('.xlsx'):
        try:
            wb = openpyxl.load_workbook(file)
            sheet = wb.active
            for row in sheet.iter_rows(min_row=2, values_only=True):
                if row[0] and row[1]:
                    full_name = str(row[0]).strip()
                    ci = str(row[1]).strip()
                    existing = Patient.query.filter_by(ci=ci).first()
                    if not existing:
                        p = Patient(
                            full_name=full_name,
                            ci=ci,
                            age=int(row[2]) if row[2] else 30,
                            gender=str(row[3]).strip() if row[3] else 'Hombre',
                            phone=str(row[4]).strip() if row[4] else '',
                            sanatorio=str(row[5]).strip() if row[5] else 'Consultorio Principal',
                            pathology_type=str(row[6]).strip() if row[6] else 'Diabetes Tipo 2',
                            evolution_years=int(row[7]) if row[7] else 0,
                            comorbidities=str(row[8]).strip() if row[8] else ''
                        )
                        db.session.add(p)
            db.session.commit()
        except Exception as e:
            print(f"Error al importar Excel: {e}")
    return redirect(url_for('index'))

@app.route('/export/excel')
@login_required
def export_excel():
    output = io.StringIO()
    writer = csv.writer(output)
    writer.writerow([
        'ID', 'Nombre', 'CI', 'Edad', 'Género', 'Teléfono', 'Sanatorio', 'Patología', 'Años Evolución', 'Comorbilidades',
        'Fecha Control', 'Peso (kg)', 'Altura (m)', 'IMC', 'Cintura (cm)', '% Grasa', 'Masa Musc', 'Grasa Visceral',
        'Glucemia', 'Insulina', 'HbA1c (%)', 'Col. Tot', 'HDL', 'LDL', 'Triglicéridos', 'PA Sist', 'PA Diast',
        'Creatinina', 'eGFR', 'Albuminuria', 'ALT/TGP', 'AST/TGO', 'HOMA-IR', 'Cintura/Altura', 'Gasto Basal', 
        'Motivo Consulta', 'Examen Físico', 'Criterio Médico', 'Conducta / Plan', 'Medicación'
    ])
    patients = Patient.query.all()
    for p in patients:
        if p.records:
            for r in p.records:
                writer.writerow([
                    p.id, p.full_name, p.ci, p.age, p.gender, p.phone, p.sanatorio, p.pathology_type, p.evolution_years, p.comorbidities,
                    r.date, r.weight, r.height, r.bmi, r.waist, r.body_fat, r.muscle_mass, r.visceral_fat,
                    r.glucose, r.insulin, r.hba1c, r.cholesterol_tot, r.hdl, r.ldl, r.triglycerides, r.bp_systolic, r.bp_diastolic,
                    r.creatinine, r.egfr, r.albuminuria, r.alt_tgp, r.ast_tgo, r.homair, r.waist_height_ratio, r.basal_metabolic_rate,
                    r.motivo_consulta, r.examen_fisico, r.notes, r.conducta_medica, r.medications
                ])
        else:
            writer.writerow([p.id, p.full_name, p.ci, p.age, p.gender, p.phone, p.sanatorio, p.pathology_type, p.evolution_years, p.comorbidities, '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', ''])
    output.seek(0)
    return Response(
        output.getvalue(),
        mimetype="text/csv",
        headers={"Content-Disposition": "attachment;filename=base_datos_investigacion_metabolica.csv"}
    )

@app.route('/report/pdf/<int:patient_id>')
@login_required
def generate_pdf(patient_id):
    patient = Patient.query.get_or_404(patient_id)
    buffer = io.BytesIO()
    p = canvas.Canvas(buffer, pagesize=letter)
    p.setFont("Helvetica-Bold", 16)
    p.drawString(50, 750, "METABOLISYS - INFORME CLÍNICO METABÓLICO")
    p.setFont("Helvetica", 10)
    p.drawString(50, 735, f"Sanatorio: {patient.sanatorio} | Patología: {patient.pathology_type}")
    p.line(50, 725, 560, 725)
    p.setFont("Helvetica-Bold", 12)
    p.drawString(50, 700, f"Paciente: {patient.full_name}")
    p.setFont("Helvetica", 10)
    p.drawString(50, 685, f"CI: {patient.ci} | Edad: {patient.age} años | Género: {patient.gender} | Años Evolución: {patient.evolution_years}")
    y = 650
    p.setFont("Helvetica-Bold", 11)
    p.drawString(50, y, "Evolución y Notas Clínicas Recientes:")
    y -= 20
    for r in patient.records[-3:]:
        p.setFont("Helvetica-Bold", 9)
        p.drawString(50, y, f"Fecha: {r.date} | Peso: {r.weight}kg | IMC: {r.bmi} | HbA1c: {r.hba1c}% | Gluc: {r.glucose}mg/dL")
        y -= 15
        p.setFont("Helvetica", 9)
        if r.motivo_consulta:
            p.drawString(60, y, f"• Motivo: {r.motivo_consulta}")
            y -= 14
        if r.conducta_medica:
            p.drawString(60, y, f"• Conducta / Plan: {r.conducta_medica}")
            y -= 14
        y -= 10
        if y < 50:
            p.showPage()
            y = 750
    p.showPage()
    p.save()
    buffer.seek(0)
    return send_file(buffer, as_attachment=True, download_name=f"Informe_Metabolisys_{patient.full_name}.pdf", mimetype='application/pdf')

@app.route('/report/recipe/<int:patient_id>')
@login_required
def generate_recipe(patient_id):
    patient = Patient.query.get_or_404(patient_id)
    buffer = io.BytesIO()
    p = canvas.Canvas(buffer, pagesize=letter)
    p.setFont("Helvetica-Bold", 18)
    p.drawString(50, 750, "METABOLISYS - RECETA MÉDICA")
    p.setFont("Helvetica", 10)
    p.drawString(50, 735, f"Atención en: {patient.sanatorio} | Tel: {patient.phone or 'N/D'}")
    p.line(50, 725, 560, 725)
    p.setFont("Helvetica-Bold", 12)
    p.drawString(50, 700, f"Paciente: {patient.full_name}")
    p.setFont("Helvetica", 10)
    p.drawString(380, 700, f"CI: {patient.ci} | Edad: {patient.age} años")
    p.drawString(50, 680, f"Fecha: _______________________")
    p.line(50, 670, 560, 670)
    p.setFont("Helvetica-Bold", 14)
    p.drawString(50, 630, "RP/")
    p.setFont("Helvetica", 11)
    y = 600
    last_record = patient.records[-1] if patient.records else None
    meds_text = last_record.medications if last_record and last_record.medications else "1. ___________________________________________________\n2. ___________________________________________________"
    for line in meds_text.split('\n'):
        p.drawString(80, y, line)
        y -= 25
    p.line(350, 150, 530, 150)
    p.setFont("Helvetica", 10)
    p.drawString(380, 135, "Firma y Sello Médico")
    p.showPage()
    p.save()
    buffer.seek(0)
    return send_file(buffer, as_attachment=True, download_name=f"Receta_{patient.full_name}.pdf", mimetype='application/pdf')

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
