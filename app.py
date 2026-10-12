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
    password_hash = db.Column(db.Text, nullable=False)

class Patient(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    full_name = db.Column(db.String(100), nullable=False)
    ci = db.Column(db.String(20), unique=True, nullable=False)
    birth_date = db.Column(db.String(20))
    age = db.Column(db.Integer, nullable=False)
    gender = db.Column(db.String(10), nullable=False)
    phone = db.Column(db.String(20))
    sanatorio = db.Column(db.String(100), nullable=False)
    procedencia = db.Column(db.String(100))
    profession = db.Column(db.String(100))
    pathology_type = db.Column(db.String(100))
    evolution_years = db.Column(db.Integer)
    antecedentes_familiares = db.Column(db.Text)
    antecedentes_personales = db.Column(db.Text)
    comorbidities = db.Column(db.String(200))
    tratamiento_base = db.Column(db.Text)
    records = db.relationship('MedicalRecord', backref='patient', lazy=True, cascade="all, delete-orphan")

class MedicalRecord(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    patient_id = db.Column(db.Integer, db.ForeignKey('patient.id'), nullable=False)
    date = db.Column(db.String(20), nullable=False)
    
    # Antropometría (Hoja Controles)
    height = db.Column(db.Float)
    weight = db.Column(db.Float)
    body_fat_kg = db.Column(db.Float)
    body_fat_pct = db.Column(db.Float)
    visceral_fat = db.Column(db.Float)
    muscle_mass_kg = db.Column(db.Float)
    muscle_mass_pct = db.Column(db.Float)
    basal_metabolism = db.Column(db.Float)
    waist = db.Column(db.Float) # Circunferencia de abdomen
    
    # Laboratorios (Hoja Laboratorios)
    glucose = db.Column(db.Float)
    insulin = db.Column(db.Float)
    hba1c = db.Column(db.Float)
    triglycerides = db.Column(db.Float)
    cholesterol_tot = db.Column(db.Float)
    hdl = db.Column(db.Float)
    ldl = db.Column(db.Float)
    vldl = db.Column(db.Float)
    uric_acid = db.Column(db.Float)
    got = db.Column(db.Float)
    gpt = db.Column(db.Float)
    fa = db.Column(db.Float)
    bili_tot = db.Column(db.Float)
    bili_dir = db.Column(db.Float)
    bili_ind = db.Column(db.Float)
    urea = db.Column(db.Float)
    creatinine = db.Column(db.Float)
    tsh = db.Column(db.Float)
    t4 = db.Column(db.Float)
    t3 = db.Column(db.Float)
    anti_tpo = db.Column(db.Float)
    vit_d = db.Column(db.Float)
    vit_b12 = db.Column(db.Float)
    calcium = db.Column(db.Float)
    magnesium = db.Column(db.Float)
    phosphorus = db.Column(db.Float)
    chlorine = db.Column(db.Float)
    
    # Estudios
    ecografia_abdominal = db.Column(db.Text)
    ecografia_tiroidea = db.Column(db.Text)
    otros_estudios = db.Column(db.Text)
    
    # Indicaciones y Tratamiento (Hoja Indicaciones)
    indicaciones_tratamiento = db.Column(db.Text)
    medications = db.Column(db.Text)
    proximo_control = db.Column(db.String(30))
    notes = db.Column(db.Text)
    conducta_medica = db.Column(db.Text)

    @property
    def bmi(self):
        if self.weight and self.height and self.height > 0:
            h_m = self.height / 100.0 if self.height > 3 else self.height
            return round(self.weight / (h_m ** 2), 2)
        return 0

    @property
    def homair(self):
        if self.glucose and self.insulin:
            return round((self.glucose * self.insulin) / 405, 2)
        return None

    @property
    def waist_height_ratio(self):
        if self.waist and self.height and self.height > 0:
            h_cm = self.height if self.height > 3 else self.height * 100
            return round(self.waist / h_cm, 2)
        return None

with app.app_context():
    db.create_all()
    admin_user = User.query.filter_by(username='admin').first()
    if not admin_user:
        hashed_pw = generate_password_hash('123', method='pbkdf2:sha256')
        admin = User(username='admin', password_hash=hashed_pw)
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
            birth_date=request.form.get('birth_date'),
            age=int(request.form['age'] or 0),
            gender=request.form['gender'],
            phone=request.form.get('phone'),
            sanatorio=request.form['sanatorio'],
            procedencia=request.form.get('procedencia'),
            profession=request.form.get('profession'),
            pathology_type=request.form.get('pathology_type'),
            evolution_years=int(request.form.get('evolution_years') or 0),
            antecedentes_familiares=request.form.get('antecedentes_familiares'),
            antecedentes_personales=request.form.get('antecedentes_personales'),
            comorbidities=request.form.get('comorbidities'),
            tratamiento_base=request.form.get('tratamiento_base')
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
    hba1c_data = [r.hba1c for r in patient.records if r.hba1c]
    weight_data = [r.weight for r in patient.records if r.weight]

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
            height=float(request.form.get('height') or 0),
            weight=float(request.form.get('weight') or 0),
            body_fat_kg=float(request.form.get('body_fat_kg') or 0),
            body_fat_pct=float(request.form.get('body_fat_pct') or 0),
            visceral_fat=float(request.form.get('visceral_fat') or 0),
            muscle_mass_kg=float(request.form.get('muscle_mass_kg') or 0),
            muscle_mass_pct=float(request.form.get('muscle_mass_pct') or 0),
            basal_metabolism=float(request.form.get('basal_metabolism') or 0),
            waist=float(request.form.get('waist') or 0),
            
            glucose=float(request.form.get('glucose') or 0),
            insulin=float(request.form.get('insulin') or 0),
            hba1c=float(request.form.get('hba1c') or 0),
            triglycerides=float(request.form.get('triglycerides') or 0),
            cholesterol_tot=float(request.form.get('cholesterol_tot') or 0),
            hdl=float(request.form.get('hdl') or 0),
            ldl=float(request.form.get('ldl') or 0),
            vldl=float(request.form.get('vldl') or 0),
            uric_acid=float(request.form.get('uric_acid') or 0),
            got=float(request.form.get('got') or 0),
            gpt=float(request.form.get('gpt') or 0),
            fa=float(request.form.get('fa') or 0),
            bili_tot=float(request.form.get('bili_tot') or 0),
            bili_dir=float(request.form.get('bili_dir') or 0),
            bili_ind=float(request.form.get('bili_ind') or 0),
            urea=float(request.form.get('urea') or 0),
            creatinine=float(request.form.get('creatinine') or 0),
            tsh=float(request.form.get('tsh') or 0),
            t4=float(request.form.get('t4') or 0),
            t3=float(request.form.get('t3') or 0),
            anti_tpo=float(request.form.get('anti_tpo') or 0),
            vit_d=float(request.form.get('vit_d') or 0),
            vit_b12=float(request.form.get('vit_b12') or 0),
            calcium=float(request.form.get('calcium') or 0),
            magnesium=float(request.form.get('magnesium') or 0),
            phosphorus=float(request.form.get('phosphorus') or 0),
            chlorine=float(request.form.get('chlorine') or 0),
            
            ecografia_abdominal=request.form.get('ecografia_abdominal'),
            ecografia_tiroidea=request.form.get('ecografia_tiroidea'),
            otros_estudios=request.form.get('otros_estudios'),
            
            indicaciones_tratamiento=request.form.get('indicaciones_tratamiento'),
            medications=request.form.get('medications'),
            proximo_control=request.form.get('proximo_control'),
            notes=request.form.get('notes'),
            conducta_medica=request.form.get('conducta_medica')
        )
        db.session.add(record)
        db.session.commit()
        return redirect(url_for('patient_detail', id=patient.id))
    return render_template('new_record.html', patient=patient)

@app.route('/export/excel')
@login_required
def export_excel():
    output = io.StringIO()
    writer = csv.writer(output)
    writer.writerow([
        'ID', 'Nombre', 'CI', 'Edad', 'Género', 'Teléfono', 'Sanatorio', 'Procedencia', 'Profesión', 'Patología', 
        'Años Evolución', 'Antecedentes Familiares', 'Antecedentes Personales', 'Comorbilidades', 'Tratamiento Base',
        'Fecha Control', 'Altura (cm)', 'Peso (kg)', 'IMC', 'Grasa Corp (kg)', 'Grasa Corp (%)', 'Grasa Visceral', 
        'Músculo (kg)', 'Músculo (%)', 'Metabolismo Basal', 'Cintura (cm)', 'Glicemia', 'Insulina', 'HbA1c (%)', 
        'Triglicéridos', 'Col. Tot', 'HDL', 'LDL', 'VLDL', 'Ácido Úrico', 'GOT', 'GPT', 'FA', 'Bili Tot', 'Bili Dir', 
        'Bili Ind', 'Urea', 'Creatinina', 'TSH', 'T4', 'T3', 'Anti-TPO', 'Vit D', 'Vit B12', 'Calcio', 'Magnesio', 
        'Fósforo', 'Cloro', 'Eco Abdominal', 'Eco Tiroidea', 'Otros Estudios', 'Indicaciones / Tratamiento', 'Medicación', 
        'Próximo Control', 'HOMA-IR', 'Cintura/Altura', 'Conducta Médica'
    ])
    patients = Patient.query.all()
    for p in patients:
        if p.records:
            for r in p.records:
                writer.writerow([
                    p.id, p.full_name, p.ci, p.age, p.gender, p.phone, p.sanatorio, p.procedencia, p.profession, p.pathology_type,
                    p.evolution_years, p.antecedentes_familiares, p.antecedentes_personales, p.comorbidities, p.tratamiento_base,
                    r.date, r.height, r.weight, r.bmi, r.body_fat_kg, r.body_fat_pct, r.visceral_fat, r.muscle_mass_kg, r.muscle_mass_pct,
                    r.basal_metabolism, r.waist, r.glucose, r.insulin, r.hba1c, r.triglycerides, r.cholesterol_tot, r.hdl, r.ldl,
                    r.vldl, r.uric_acid, r.got, r.gpt, r.fa, r.bili_tot, r.bili_dir, r.bili_ind, r.urea, r.creatinine,
                    r.tsh, r.t4, r.t3, r.anti_tpo, r.vit_d, r.vit_b12, r.calcium, r.magnesium, r.phosphorus, r.chlorine,
                    r.ecografia_abdominal, r.ecografia_tiroidea, r.otros_estudios, r.indicaciones_tratamiento, r.medications,
                    r.proximo_control, r.homair, r.waist_height_ratio, r.conducta_medica
                ])
        else:
            writer.writerow([p.id, p.full_name, p.ci, p.age, p.gender, p.phone, p.sanatorio, p.procedencia, p.profession, p.pathology_type, p.evolution_years, p.antecedentes_familiares, p.antecedentes_personales, p.comorbidities, p.tratamiento_base, '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', ''])
    output.seek(0)
    return Response(
        output.getvalue(),
        mimetype="text/csv",
        headers={"Content-Disposition": "attachment;filename=matriz_metabolica_pacientes.csv"}
    )

@app.route('/report/pdf/<int:patient_id>')
@login_required
def generate_pdf(patient_id):
    patient = Patient.query.get_or_404(patient_id)
    buffer = io.BytesIO()
    p = canvas.Canvas(buffer, pagesize=letter)
    
    # Encabezado
    p.setFont("Helvetica-Bold", 14)
    p.drawString(50, 750, "DRA. SOFIA FERNANDEZ - DIABETES Y OBESIDAD")
    p.setFont("Helvetica", 9)
    p.drawString(50, 735, f"Sanatorio: {patient.sanatorio} | Tel: {patient.phone or 'N/D'}")
    p.line(50, 725, 560, 725)
    
    # Datos del Paciente
    p.setFont("Helvetica-Bold", 11)
    p.drawString(50, 705, f"INFORME CLÍNICO DE PACIENTE: {patient.full_name}")
    p.setFont("Helvetica", 10)
    p.drawString(50, 690, f"CI: {patient.ci} | Edad: {patient.age} años | Género: {patient.gender} | Procedencia: {patient.procedencia or 'N/D'}")
    
    y = 665
    
    if patient.records:
        first_r = patient.records[0]
        last_r = patient.records[-1]
        
        w_init = first_r.weight or 0
        w_last = last_r.weight or 0
        w_diff = round(w_init - w_last, 2)
        
        f_init = first_r.body_fat_kg or 0
        f_last = last_r.body_fat_kg or 0
        f_diff = round(f_init - f_last, 2)
        
        m_init = first_r.muscle_mass_kg or 0
        m_last = last_r.muscle_mass_kg or 0
        m_diff = round(m_last - m_init, 2)
        
        c_init = first_r.waist or 0
        c_last = last_r.waist or 0
        c_diff = round(c_init - c_last, 2)
        
        # Cuadro de Resumen Desde Inicio al Último Control
        p.setFont("Helvetica-Bold", 10)
        p.drawString(50, y, "RESUMEN DESDE INICIO AL ÚLTIMO CONTROL")
        y -= 15
        p.setFont("Helvetica", 9)
        p.drawString(50, y, f"• Peso Inicial: {w_init} kg  |  Peso Último: {w_last} kg  |  Peso Total Perdido: {w_diff} kg")
        y -= 14
        p.drawString(50, y, f"• Grasa Inicial: {f_init} kg  |  Grasa Última: {f_last} kg  |  Grasa Total Perdida: {f_diff} kg")
        y -= 14
        p.drawString(50, y, f"• Músculo Inicial: {m_init} kg  |  Músculo Último: {m_last} kg  |  Variación Músculo: {m_diff} kg")
        y -= 14
        p.drawString(50, y, f"• Cintura Inicial: {c_init} cm  |  Cintura Última: {c_last} cm  |  Cintura Perdida: {c_diff} cm")
        y -= 25
        
        # Historial de Controles
        p.setFont("Helvetica-Bold", 10)
        p.drawString(50, y, "EVOLUCIÓN POR FECHAS (CONTROLES ANTROPOMÉTRICOS)")
        y -= 15
        
        p.setFont("Helvetica-Bold", 8)
        p.drawString(50, y, "FECHA")
        p.drawString(110, y, "ALTURA")
        p.drawString(170, y, "PESO")
        p.drawString(230, y, "GRASA (kg)")
        p.drawString(300, y, "MÚSCULO")
        p.drawString(370, y, "IMC")
        p.drawString(430, y, "CINTURA")
        y -= 12
        p.line(50, y+3, 540, y+3)
        y -= 12
        
        p.setFont("Helvetica", 8)
        for r in patient.records:
            if y < 150:
                p.showPage()
                y = 750
            p.drawString(50, y, str(r.date))
            p.drawString(110, y, f"{r.height or '-'} cm")
            p.drawString(170, y, f"{r.weight or '-'} kg")
            p.drawString(230, y, f"{r.body_fat_kg or '-'} kg")
            p.drawString(300, y, f"{r.muscle_mass_kg or '-'} kg")
            p.drawString(370, y, str(r.bmi))
            p.drawString(430, y, f"{r.waist or '-'} cm")
            y -= 16
        
        y -= 15
        if y < 180:
            p.showPage()
            y = 750
            
        p.setFont("Helvetica-Bold", 10)
        p.drawString(50, y, "INDICACIONES Y TRATAMIENTO ACTUAL:")
        y -= 16
        p.setFont("Helvetica", 9)
        
        meds = last_r.medications or "Sin medicación registrada"
        for line in meds.split('\n'):
            p.drawString(60, y, f"• {line}")
            y -= 14
            
        inds = last_r.indicaciones_tratamiento or "Sin indicaciones adicionales"
        for line in inds.split('\n'):
            p.drawString(60, y, f"• {line}")
            y -= 14
    else:
        p.setFont("Helvetica", 10)
        p.drawString(50, y, "No hay controles registrados todavía para este paciente.")

    p.showPage()
    p.save()
    buffer.seek(0)
    return send_file(buffer, as_attachment=True, download_name=f"Informe_{patient.full_name}.pdf", mimetype='application/pdf')

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
