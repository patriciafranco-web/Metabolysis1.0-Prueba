from flask import Flask, render_template, request, redirect, url_for, Response, send_file, session
from flask_sqlalchemy import SQLAlchemy
from werkzeug.security import generate_password_hash, check_password_hash
from werkzeug.utils import secure_filename
from functools import wraps
import io
import csv
import os
import openpyxl
from datetime import date
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image as RLImage
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

app = Flask(__name__)
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///metabolisys_pro.db'
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
app.secret_key = 'metabolisys_secreto_seguro_2026'

UPLOAD_FOLDER = 'static/uploads'
ALLOWED_EXTENSIONS = {'png', 'jpg', 'jpeg', 'webp'}
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def format_text(text):
    if not text:
        return ""
    lines = text.split('\n')
    formatted_lines = [" ".join([w.capitalize() for w in line.split()]) for line in lines if line.strip()]
    return "\n".join(formatted_lines)

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
    
    seguro_medico = db.Column(db.String(100))
    fecha_inicio_tratamiento = db.Column(db.String(50))

    antecedentes_familiares = db.Column(db.Text)
    antecedentes_personales = db.Column(db.Text)
    comorbidities = db.Column(db.String(200))
    tratamiento_base = db.Column(db.Text)
    historial_cirugias = db.Column(db.Text)
    alergia_farmacos = db.Column(db.Text)
    cant_hijos = db.Column(db.String(50))
    menarca = db.Column(db.String(50))
    menstruacion_regular = db.Column(db.String(50))
    anticonceptivos = db.Column(db.String(50))
    checklist_glp1_notas = db.Column(db.Text)
    calidad_sueno = db.Column(db.String(100))
    actividad_fisica = db.Column(db.String(100))
    fuma_alcohol = db.Column(db.String(100))
    veces_come = db.Column(db.String(100))
    consume_gaseosa_jugo = db.Column(db.String(100))

    records = db.relationship('MedicalRecord', backref='patient', lazy=True, cascade="all, delete-orphan")

class MedicalRecord(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    patient_id = db.Column(db.Integer, db.ForeignKey('patient.id'), nullable=False)
    date = db.Column(db.String(20), nullable=False)
    
    marca_balanza = db.Column(db.String(50))
    pa = db.Column(db.String(20))
    fc = db.Column(db.Integer)
    
    height = db.Column(db.Float)
    weight = db.Column(db.Float)
    body_fat_kg = db.Column(db.Float)
    body_fat_pct = db.Column(db.Float)
    visceral_fat = db.Column(db.Float)
    muscle_mass_kg = db.Column(db.Float)
    muscle_mass_pct = db.Column(db.Float)
    basal_metabolism = db.Column(db.Float)
    waist = db.Column(db.Float)
    
    masa_osea = db.Column(db.Float)
    agua_corporal_pct = db.Column(db.Float)
    grasa_subcutanea_pct = db.Column(db.Float)
    proteina_pct = db.Column(db.Float)
    
    weight_lost = db.Column(db.Float, default=0.0)
    fat_lost = db.Column(db.Float, default=0.0)
    muscle_delta = db.Column(db.Float, default=0.0)
    waist_lost = db.Column(db.Float, default=0.0)
    
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
    
    ecografia_abdominal = db.Column(db.Text)
    ecografia_tiroidea = db.Column(db.Text)
    otros_estudios = db.Column(db.Text)
    
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

class PatientPhoto(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    patient_id = db.Column(db.Integer, db.ForeignKey('patient.id'), nullable=False)
    record_id = db.Column(db.Integer, db.ForeignKey('medical_record.id'), nullable=True)
    filename = db.Column(db.String(200), nullable=False)
    category = db.Column(db.String(50), nullable=False)
    date = db.Column(db.String(20), nullable=False)
    notes = db.Column(db.Text)
    
    record = db.relationship('MedicalRecord', backref='photos', lazy=True)

Patient.photos = db.relationship('PatientPhoto', backref='patient', lazy=True, cascade="all, delete-orphan")

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
            full_name=format_text(request.form['full_name']),
            ci=request.form['ci'].strip(),
            birth_date=request.form.get('birth_date'),
            age=int(request.form['age'] or 0),
            gender=request.form['gender'],
            phone=request.form.get('phone'),
            sanatorio=format_text(request.form['sanatorio']),
            procedencia=format_text(request.form.get('procedencia')),
            profession=format_text(request.form.get('profession')),
            pathology_type=format_text(request.form.get('pathology_type')),
            evolution_years=int(request.form.get('evolution_years') or 0),
            seguro_medico=format_text(request.form.get('seguro_medico')),
            fecha_inicio_tratamiento=format_text(request.form.get('fecha_inicio_tratamiento')),
            antecedentes_familiares=format_text(request.form.get('antecedentes_familiares')),
            antecedentes_personales=format_text(request.form.get('antecedentes_personales')),
            comorbidities=format_text(request.form.get('comorbidities')),
            tratamiento_base=format_text(request.form.get('tratamiento_base')),
            historial_cirugias=format_text(request.form.get('historial_cirugias')),
            alergia_farmacos=format_text(request.form.get('alergia_farmacos')),
            cant_hijos=format_text(request.form.get('cant_hijos')),
            menarca=format_text(request.form.get('menarca')),
            menstruacion_regular=format_text(request.form.get('menstruacion_regular')),
            anticonceptivos=format_text(request.form.get('anticonceptivos')),
            checklist_glp1_notas=format_text(request.form.get('checklist_glp1_notas')),
            calidad_sueno=format_text(request.form.get('calidad_sueno')),
            actividad_fisica=format_text(request.form.get('actividad_fisica')),
            fuma_alcohol=format_text(request.form.get('fuma_alcohol')),
            veces_come=format_text(request.form.get('veces_come')),
            consume_gaseosa_jugo=format_text(request.form.get('consume_gaseosa_jugo'))
        )
        db.session.add(patient)
        db.session.commit()
        return redirect(url_for('index'))
    return render_template('new_patient.html')

@app.route('/patient/<int:id>')
@login_required
def patient_detail(id):
    patient = Patient.query.get_or_404(id)
    valid_records = [r for r in patient.records if r.weight is not None and r.weight > 0]
    dates = [r.date for r in valid_records]
    hba1c_data = [r.hba1c for r in valid_records if r.hba1c]
    weight_data = [r.weight for r in valid_records if r.weight]

    return render_template(
        'patient_detail.html', 
        patient=patient, 
        valid_records=valid_records,
        dates=dates, 
        hba1c_data=hba1c_data, 
        weight_data=weight_data
    )

@app.route('/record/new/<int:patient_id>', methods=['GET', 'POST'])
@login_required
def new_record(patient_id):
    patient = Patient.query.get_or_404(patient_id)
    if request.method == 'POST':
        valid_records = [r for r in patient.records if r.weight is not None and r.weight > 0]
        first_record = valid_records[0] if valid_records else None
        
        w_curr = float(request.form.get('weight') or 0)
        f_curr = float(request.form.get('body_fat_kg') or 0)
        m_curr = float(request.form.get('muscle_mass_kg') or 0)
        wst_curr = float(request.form.get('waist') or 0)
        
        w_init = first_record.weight if first_record and first_record.weight else w_curr
        f_init = first_record.body_fat_kg if first_record and first_record.body_fat_kg else f_curr
        m_init = first_record.muscle_mass_kg if first_record and first_record.muscle_mass_kg else m_curr
        wst_init = first_record.waist if first_record and first_record.waist else wst_curr

        record = MedicalRecord(
            patient_id=patient.id,
            date=request.form['date'],
            marca_balanza=request.form.get('marca_balanza'),
            pa=request.form.get('pa'),
            fc=int(request.form.get('fc') or 0),
            height=float(request.form.get('height') or 0),
            weight=w_curr,
            body_fat_kg=f_curr,
            body_fat_pct=float(request.form.get('body_fat_pct') or 0),
            visceral_fat=float(request.form.get('visceral_fat') or 0),
            muscle_mass_kg=m_curr,
            muscle_mass_pct=float(request.form.get('muscle_mass_pct') or 0),
            basal_metabolism=float(request.form.get('basal_metabolism') or 0),
            waist=wst_curr,
            
            masa_osea=float(request.form.get('masa_osea') or 0),
            agua_corporal_pct=float(request.form.get('agua_corporal_pct') or 0),
            grasa_subcutanea_pct=float(request.form.get('grasa_subcutanea_pct') or 0),
            proteina_pct=float(request.form.get('proteina_pct') or 0),
            
            weight_lost=round(w_init - w_curr, 2) if first_record else 0.0,
            fat_lost=round(f_init - f_curr, 2) if first_record else 0.0,
            muscle_delta=round(m_curr - m_init, 2) if first_record else 0.0,
            waist_lost=round(wst_init - wst_curr, 2) if first_record else 0.0,
            
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
            
            indicaciones_tratamiento=format_text(request.form.get('indicaciones_tratamiento')),
            medications=format_text(request.form.get('medications')),
            proximo_control=request.form.get('proximo_control'),
            notes=request.form.get('notes'),
            conducta_medica=format_text(request.form.get('conducta_medica'))
        )
        db.session.add(record)
        db.session.commit()

        if 'photo' in request.files:
            file = request.files['photo']
            if file and allowed_file(file.filename):
                filename = secure_filename(f"p_{patient.id}_r_{record.id}_{file.filename}")
                file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
                
                photo = PatientPhoto(
                    patient_id=patient.id,
                    record_id=record.id,
                    filename=filename,
                    category=request.form.get('photo_category', 'Evolución'),
                    date=request.form['date'],
                    notes=request.form.get('photo_notes', '')
                )
                db.session.add(photo)
                db.session.commit()

        return redirect(url_for('patient_detail', id=patient.id))
    return render_template('new_record.html', patient=patient)

@app.route('/treatment/edit/<int:patient_id>', methods=['GET', 'POST'])
@login_required
def edit_treatment(patient_id):
    patient = Patient.query.get_or_404(patient_id)
    last_record = MedicalRecord.query.filter_by(patient_id=patient.id).order_by(MedicalRecord.id.desc()).first()
    
    if request.method == 'POST':
        today_str = date.today().strftime('%Y-%m-%d')
        if not last_record:
            last_record = MedicalRecord(patient_id=patient.id, date=today_str, weight=None, height=None)
            db.session.add(last_record)
            db.session.commit()
            
        last_record.indicaciones_tratamiento = format_text(request.form.get('indicaciones_tratamiento'))
        last_record.medications = format_text(request.form.get('medications'))
        last_record.proximo_control = request.form.get('proximo_control')
        last_record.conducta_medica = format_text(request.form.get('conducta_medica'))
        db.session.commit()

        if 'photo' in request.files:
            file = request.files['photo']
            if file and allowed_file(file.filename):
                filename = secure_filename(f"p_{patient.id}_r_{last_record.id}_{file.filename}")
                file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
                
                photo = PatientPhoto(
                    patient_id=patient.id,
                    record_id=last_record.id,
                    filename=filename,
                    category=request.form.get('photo_category', 'Evolución'),
                    date=last_record.date,
                    notes=request.form.get('photo_notes', '')
                )
                db.session.add(photo)
                db.session.commit()

        return redirect(url_for('patient_detail', id=patient.id))
        
    return render_template('edit_treatment.html', patient=patient, record=last_record)

@app.route('/export/excel')
@login_required
def export_excel():
    output = io.StringIO()
    writer = csv.writer(output)
    writer.writerow([
        'ID', 'Nombre', 'CI', 'Edad', 'Género', 'Teléfono', 'Sanatorio', 'Procedencia', 'Profesión', 'Patología', 
        'Seguro Médico', 'Inicio Tratamiento', 'Años Evolución', 'Antecedentes Familiares', 'Antecedentes Personales', 'Comorbilidades', 'Tratamiento Base',
        'Fecha Control', 'Balanza', 'PA', 'FC', 'Altura (cm)', 'Peso (kg)', 'Peso Perdido (kg)', 'IMC', 'Grasa Corp (kg)', 'Grasa Perdida (kg)', 'Grasa Corp (%)', 'Grasa Visceral', 
        'Músculo (kg)', 'Músculo Delta (kg)', 'Músculo (%)', 'Metabolismo Basal', 'Masa Ósea', 'Agua Corporal (%)', 'Grasa Subcutánea (%)', 'Proteína (%)', 'Cintura (cm)', 'Cintura Perdida (cm)', 'Glicemia', 'Insulina', 'HbA1c (%)', 
        'Triglicéridos', 'Col. Tot', 'HDL', 'LDL', 'VLDL', 'Ácido Úrico', 'GOT', 'GPT', 'FA', 'Bili Tot', 'Bili Dir', 
        'Bili Ind', 'Urea', 'Creatinina', 'TSH', 'T4', 'T3', 'Anti-TPO', 'Vit D', 'Vit B12', 'Calcio', 'Magnesio', 
        'Fósforo', 'Cloro', 'Eco Abdominal', 'Eco Tiroidea', 'Otros Estudios', 'Indicaciones / Tratamiento', 'Medicación', 
        'Próximo Control', 'HOMA-IR', 'Cintura/Altura', 'Conducta Médica'
    ])
    patients = Patient.query.all()
    for p in patients:
        valid_recs = [r for r in p.records if r.weight is not None and r.weight > 0]
        if valid_recs:
            for r in valid_recs:
                writer.writerow([
                    p.id, p.full_name, p.ci, p.age, p.gender, p.phone, p.sanatorio, p.procedencia, p.profession, p.pathology_type,
                    p.seguro_medico, p.fecha_inicio_tratamiento, p.evolution_years, p.antecedentes_familiares, p.antecedentes_personales, p.comorbidities, p.tratamiento_base,
                    r.date, r.marca_balanza, r.pa, r.fc, r.height, r.weight, r.weight_lost, r.bmi, r.body_fat_kg, r.fat_lost, r.body_fat_pct, r.visceral_fat, r.muscle_mass_kg, r.muscle_delta, r.muscle_mass_pct,
                    r.basal_metabolism, r.masa_osea, r.agua_corporal_pct, r.grasa_subcutanea_pct, r.proteina_pct, r.waist, r.waist_lost, r.glucose, r.insulin, r.hba1c, r.triglycerides, r.cholesterol_tot, r.hdl, r.ldl,
                    r.vldl, r.uric_acid, r.got, r.gpt, r.fa, r.bili_tot, r.bili_dir, r.bili_ind, r.urea, r.creatinine,
                    r.tsh, r.t4, r.t3, r.anti_tpo, r.vit_d, r.vit_b12, r.calcium, r.magnesium, r.phosphorus, r.chlorine,
                    r.ecografia_abdominal, r.ecografia_tiroidea, r.otros_estudios, r.indicaciones_tratamiento, r.medications,
                    r.proximo_control, r.homair, r.waist_height_ratio, r.conducta_medica
                ])
        else:
            writer.writerow([p.id, p.full_name, p.ci, p.age, p.gender, p.phone, p.sanatorio, p.procedencia, p.profession, p.pathology_type, p.seguro_medico, p.fecha_inicio_tratamiento, p.evolution_years, p.antecedentes_familiares, p.antecedentes_personales, p.comorbidities, p.tratamiento_base, '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', ''])
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
    doc = SimpleDocTemplate(buffer, pagesize=letter, rightMargin=20, leftMargin=20, topMargin=20, bottomMargin=20)
    story = []
    
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        'DocTitle', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=14, alignment=1, spaceAfter=15
    )
    section_style = ParagraphStyle(
        'SectionHeading', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=10, alignment=1, spaceAfter=8, spaceBefore=10
    )
    body_style = ParagraphStyle(
        'DocBody', parent=styles['Normal'], fontName='Helvetica', fontSize=9, leading=13
    )
    footer_style = ParagraphStyle(
        'DocFooter', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=9, alignment=1, textColor=colors.HexColor('#006666')
    )
    
    logo_path = os.path.join('static', 'logo.png')
    if os.path.exists(logo_path):
        story.append(RLImage(logo_path, width=160, height=58))
        story.append(Spacer(1, 10))
    
    valid_records = [r for r in patient.records if r.weight is not None and r.weight > 0]
    last_r = valid_records[-1] if valid_records else (patient.records[-1] if patient.records else None)
    report_date = last_r.date if last_r else "N/D"
    
    story.append(Paragraph(f"{patient.full_name.upper()} &nbsp;&nbsp;&nbsp;&nbsp; {report_date}", title_style))
    story.append(Spacer(1, 5))
    
    if valid_records:
        table_data = [[
            'FECHA', 'BAL.', 'PA / FC', 'ALT.\n(cm)', 'PESO\n(kg)', 'GRASA\n(kg)', 'GRASA\n(%)', 'VISC.', 'MÚSC.\n(kg)', 'MÚSC.\n(%)', 'METAB.\n(kcal)', 'IMC', 'CA\n(cm)'
        ]]
        for r in valid_records:
            pa_fc = f"{r.pa or '-'} / {r.fc or '-'}"
            table_data.append([
                str(r.date), str(r.marca_balanza or '-'), pa_fc, str(r.height or '-'), str(r.weight or '-'), str(r.body_fat_kg or '-'),
                str(r.body_fat_pct or '-'), str(r.visceral_fat or '-'), str(r.muscle_mass_kg or '-'),
                str(r.muscle_mass_pct or '-'), str(r.basal_metabolism or '-'), str(r.bmi or '-'), str(r.waist or '-')
            ])
        
        t = Table(table_data, colWidths=[55, 45, 50, 35, 35, 38, 35, 35, 38, 35, 45, 30, 30])
        t.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#e8f4f4')),
            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
            ('FONTSIZE', (0,0), (-1,0), 6.5),
            ('FONTSIZE', (0,1), (-1,-1), 6.5),
            ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#cccccc')),
            ('BOTTOMPADDING', (0,0), (-1,-1), 4),
            ('TOPPADDING', (0,0), (-1,-1), 4),
        ]))
        story.append(t)
        story.append(Spacer(1, 10))
        
        first_r = valid_records[0]
        w_init = first_r.weight or 0
        w_last = last_r.weight or 0
        w_diff = round(w_init - w_last, 2)
        
        f_init = first_r.body_fat_kg or 0
        f_last = last_r.body_fat_kg or 0
        f_diff = round(f_init - f_last, 2)
        
        m_init = first_r.muscle_mass_kg or 0
        m_last = last_r.muscle_mass_kg or 0
        
        c_init = first_r.waist or 0
        c_last = last_r.waist or 0
        c_diff = round(c_init - c_last, 2)
        
        story.append(Paragraph("RESUMEN DESDE INICIO AL ÚLTIMO CONTROL", section_style))
        
        sum_data = [
            ['Peso inicial (kg)', 'Peso último (kg)', 'Peso total perdido (kg)'],
            [str(w_init), str(w_last), str(w_diff)],
            ['Grasa inicial (kg)', 'Grasa última (kg)', 'Grasa total perdida (kg)'],
            [str(f_init), str(f_last), str(f_diff)],
            ['Músculo inicial (kg)', 'Músculo último (kg)', 'Músculo total (kg)'],
            [str(m_init), str(m_last), str(round(m_last - m_init, 2))],
            ['Cintura inicial (cm)', 'Cintura última (cm)', 'Cintura total perdida (cm)'],
            [str(c_init), str(c_last), str(c_diff)]
        ]
        sum_table = Table(sum_data, colWidths=[180, 180, 180])
        sum_table.setStyle(TableStyle([
            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
            ('FONTNAME', (0,2), (-1,2), 'Helvetica-Bold'),
            ('FONTNAME', (0,4), (-1,4), 'Helvetica-Bold'),
            ('FONTNAME', (0,6), (-1,6), 'Helvetica-Bold'),
            ('FONTSIZE', (0,0), (-1,-1), 8),
            ('TEXTCOLOR', (0,0), (-1,0), colors.HexColor('#006666')),
            ('TEXTCOLOR', (0,2), (-1,2), colors.HexColor('#006666')),
            ('TEXTCOLOR', (0,4), (-1,4), colors.HexColor('#006666')),
            ('TEXTCOLOR', (0,6), (-1,6), colors.HexColor('#006666')),
            ('GRID', (0,0), (-1,-1), 0.75, colors.HexColor('#006666')),
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#f0f8f8')),
            ('BACKGROUND', (0,2), (-1,2), colors.HexColor('#f0f8f8')),
            ('BACKGROUND', (0,4), (-1,4), colors.HexColor('#f0f8f8')),
            ('BACKGROUND', (0,6), (-1,6), colors.HexColor('#f0f8f8')),
            ('BOTTOMPADDING', (0,0), (-1,-1), 4),
            ('TOPPADDING', (0,0), (-1,-1), 4),
        ]))
        story.append(sum_table)
        story.append(Spacer(1, 15))
        
        story.append(Paragraph("<b>INDICACIONES Y PLAN TERAPÉUTICO</b>", section_style))
        story.append(Spacer(1, 5))
        
        meds = last_r.medications or ""
        inds = last_r.indicaciones_tratamiento or ""
        
        plan_text = ""
        if meds:
            for line in meds.split('\n'):
                if line.strip():
                    plan_text += f"• &nbsp; <b>{line.upper()}</b><br/>"
        if inds:
            for line in inds.split('\n'):
                if line.strip():
                    plan_text += f"• &nbsp; {line.upper()}<br/>"
        if last_r.proximo_control:
            plan_text += f"<br/>• &nbsp; <b>CONTROL EN {last_r.proximo_control.upper()}.</b>"
            
        if not plan_text:
            plan_text = "Sin indicaciones registradas para este control."
            
        story.append(Paragraph(plan_text, body_style))
        story.append(Spacer(1, 15))
        
        story.append(Paragraph("Instagram: @drasofiafernandez &nbsp;&nbsp;|&nbsp;&nbsp; Teléfono: 0986926714", footer_style))
        
    else:
        story.append(Paragraph("No hay controles antropométricos registrados todavía para este paciente.", body_style))
        
    doc.build(story)
    buffer.seek(0)
    return send_file(buffer, as_attachment=True, download_name=f"Informe_{patient.full_name}.pdf", mimetype='application/pdf')

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
