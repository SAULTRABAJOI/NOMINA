import os
import io
import pdfkit
import pandas as pd
from datetime import datetime, timedelta
from flask import (
    Flask, render_template, request, redirect, url_for,
    send_from_directory, Response, flash
)
from flask_sqlalchemy import SQLAlchemy
from flask_login import (
    LoginManager, UserMixin, login_user,
    login_required, logout_user, current_user
)
from collections import defaultdict, Counter

# —————————————————————————————————————————————————————————
# Configuración de la aplicación
# —————————————————————————————————————————————————————————
app = Flask(__name__)
app.config['SECRET_KEY'] = 'secret-key'
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///foco.db'
app.config['UPLOAD_FOLDER'] = os.path.join(os.getcwd(), 'uploads')
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

db = SQLAlchemy(app)
login_manager = LoginManager(app)
login_manager.login_view = 'login'

# —————————————————————————————————————————————————————————
# Modelos de datos
# —————————————————————————————————————————————————————————
class Employee(db.Model, UserMixin):
    usuario        = db.Column(db.String(50), primary_key=True)
    nombre         = db.Column(db.String(100), nullable=False)
    salario_diario = db.Column(db.Float, nullable=False)
    puesto         = db.Column(db.String(100))
    password_hash  = db.Column(db.String(128))  # implementar check_password
    is_admin       = db.Column(db.Boolean, default=False)
    is_supervisor  = db.Column(db.Boolean, default=False)
    supervisor_id  = db.Column(db.String(50), db.ForeignKey('employee.usuario'))
    agentes        = db.relationship(
                        'Employee',
                        backref=db.backref('supervisor', remote_side=[usuario])
                    )

    def check_password(self, pw):
        # aquí tu lógica de verificación de contraseña
        return True


class Attendance(db.Model):
    __tablename__ = 'attendance'
    id          = db.Column(db.Integer, primary_key=True)
    usuario     = db.Column(db.String(50), db.ForeignKey('employee.usuario'), nullable=False)
    fecha       = db.Column(db.Date, nullable=False)
    estado      = db.Column(db.String(20), nullable=False)   # A, V, MG, F, D
    supervisor  = db.Column(db.String(100))
    cartera     = db.Column(db.String(100))
    __table_args__ = (
        db.UniqueConstraint('usuario', 'fecha', name='uix_usuario_fecha'),
    )


class Deduction(db.Model):
    id          = db.Column(db.Integer, primary_key=True)
    usuario     = db.Column(db.String(50), db.ForeignKey('employee.usuario'))
    fecha       = db.Column(db.Date, nullable=False)
    tipo        = db.Column(db.String(50), nullable=False)
    monto       = db.Column(db.Float, nullable=False)
    observacion = db.Column(db.String(200))


class Bonus(db.Model):
    id          = db.Column(db.Integer, primary_key=True)
    usuario     = db.Column(db.String(50), db.ForeignKey('employee.usuario'))
    fecha       = db.Column(db.Date, nullable=False)
    tipo        = db.Column(db.String(50), nullable=False)
    monto       = db.Column(db.Float, nullable=False)
    observacion = db.Column(db.String(200))


class Payroll(db.Model):
    id                = db.Column(db.Integer, primary_key=True)
    usuario           = db.Column(db.String(50), db.ForeignKey('employee.usuario'))
    inicio            = db.Column(db.Date, nullable=False)
    fin               = db.Column(db.Date, nullable=False)
    sueldo_base       = db.Column(db.Float)
    total_bonos       = db.Column(db.Float)
    total_deducciones = db.Column(db.Float)
    neto              = db.Column(db.Float)
    employee          = db.relationship('Employee', backref='payrolls')


# —————————————————————————————————————————————————————————
# Usuario admin por defecto (opcional)
# —————————————————————————————————————————————————————————
class Admin(UserMixin):
    id = 'admin'

@login_manager.user_loader
def load_user(user_id):
    if user_id == 'admin':
        return Admin()
    return Employee.query.get(user_id)


# —————————————————————————————————————————————————————————
# Rutas de autenticación
# —————————————————————————————————————————————————————————
@app.route('/login', methods=['GET','POST'])
def login():
    if request.method=='POST':
        usr = request.form['username'].strip()
        pwd = request.form['password']
        if usr == 'admin' and pwd == 'admin123':
            login_user(Admin())
            return redirect(url_for('dashboard'))
        emp = Employee.query.get(usr)
        if emp and emp.check_password(pwd):
            login_user(emp)
            return redirect(url_for('dashboard'))
        flash("Usuario o contraseña inválidos", "danger")
    return render_template('login.html')


@app.route('/logout')
@login_required
def logout():
    logout_user()
    return redirect(url_for('login'))


# —————————————————————————————————————————————————————————
# Dashboard
# —————————————————————————————————————————————————————————
@app.route('/')
@login_required
def dashboard():
    return render_template('dashboard.html')


# —————————————————————————————————————————————————————————
# CRUD de Empleados (listado, crear, editar, borrar, subir masivo)
# —————————————————————————————————————————————————————————
@app.route('/employees')
@login_required
def list_employees():
    if current_user.is_admin:
        empleados = Employee.query.order_by(Employee.usuario).all()
    elif current_user.is_supervisor:
        empleados = current_user.agentes
    else:
        empleados = [current_user]
    return render_template('employees.html', empleados=empleados)


@app.route('/employees/create', methods=['GET','POST'])
@login_required
def create_employee():
    supervisors = Employee.query.filter_by(is_supervisor=True).all()
    if request.method=='POST':
        usr = request.form['usuario'].strip()
        if Employee.query.get(usr):
            flash(f"El usuario «{usr}» ya está registrado.", "warning")
            return redirect(url_for('create_employee'))
        e = Employee(
            usuario        = usr,
            nombre         = request.form['nombre'].strip(),
            salario_diario = float(request.form['salario_diario']),
            puesto         = request.form['puesto'].strip(),
            is_supervisor  = ('is_supervisor' in request.form),
            supervisor_id  = request.form.get('supervisor_id') or None
        )
        db.session.add(e)
        db.session.commit()
        flash('Empleado creado correctamente.', 'success')
        return redirect(url_for('list_employees'))
    return render_template('employee_form.html',
                           action='Crear',
                           supervisors=supervisors)


@app.route('/employees/<usuario>/edit', methods=['GET','POST'])
@login_required
def edit_employee(usuario):
    e = Employee.query.get_or_404(usuario)
    supervisors = Employee.query.filter_by(is_supervisor=True).all()
    if request.method=='POST':
        e.nombre         = request.form['nombre'].strip()
        e.salario_diario = float(request.form['salario_diario'])
        e.puesto         = request.form['puesto'].strip()
        e.is_supervisor  = ('is_supervisor' in request.form)
        e.supervisor_id  = request.form.get('supervisor_id') or None
        db.session.commit()
        flash('Empleado actualizado correctamente.', 'success')
        return redirect(url_for('list_employees'))
    return render_template('employee_form.html',
                           action='Editar',
                           empleado=e,
                           supervisors=supervisors)


@app.route('/employees/<usuario>/delete', methods=['POST'])
@login_required
def delete_employee(usuario):
    e = Employee.query.get_or_404(usuario)
    db.session.delete(e)
    db.session.commit()
    flash('Empleado eliminado.', 'warning')
    return redirect(url_for('list_employees'))


@app.route('/upload/employees', methods=['GET','POST'])
@login_required
def upload_employees():
    plantilla  = 'employees.xlsx'
    errores    = []; creados = 0; existentes = 0
    if request.method=='POST':
        try:
            df = pd.read_excel(request.files['file'])
        except Exception as ex:
            flash(f"Error al leer plantilla: {ex}", 'danger')
            return redirect(url_for('upload_employees'))

        cols = ['Usuario','Nombre','Salario diario','Puesto','EsSupervisor','SupervisorID']
        falt = set(cols) - set(df.columns)
        if falt:
            flash(f"Faltan columnas: {', '.join(falt)}", 'warning')
            return redirect(url_for('upload_employees'))

        for i,row in df.iterrows():
            fila = i+2
            usr  = str(row['Usuario']).strip()
            if not usr:
                errores.append(f"Fila {fila}: usuario vacío")
                continue
            if Employee.query.get(usr):
                existentes += 1
                continue
            try:
                nombre = str(row['Nombre']).strip()
                salario= float(row['Salario diario'])
                puesto = str(row['Puesto']).strip()
                is_sup = str(row['EsSupervisor']).strip().upper() in ('TRUE','1','SI','YES')
                supid  = str(row['SupervisorID']).strip() or None
                if supid and not Employee.query.get(supid):
                    raise ValueError(f"Supervisor «{supid}» no existe")
                e = Employee(
                    usuario        = usr,
                    nombre         = nombre,
                    salario_diario = salario,
                    puesto         = puesto,
                    is_supervisor  = is_sup,
                    supervisor_id  = supid
                )
                db.session.add(e)
                creados += 1
            except Exception as ex:
                errores.append(f"Fila {fila}: {ex}")

        db.session.commit()
        if creados:   flash(f"Se crearon {creados} empleados nuevos.", 'success')
        if existentes:flash(f"{existentes} ya existían y se omitieron.", 'info')
        if errores:
            flash("Se omitieron filas con error:", 'warning')
            for msg in errores:
                flash(msg, 'warning')
        return redirect(url_for('upload_employees'))

    return render_template('upload_employees.html', plantilla=plantilla)


# —————————————————————————————————————————————————————————
# Subida masiva de Asistencia/Deducciones/Bonos
# —————————————————————————————————————————————————————————
@app.route('/upload/<tipo>', methods=['GET','POST'])
@login_required
def upload(tipo):
    plantilla = f"{tipo}.xlsx"
    errores = []; cargados = 0; reemplazados = 0; detalles = []

    if request.method=='POST':
        try:
            df = pd.read_excel(request.files['file'])
        except Exception as ex:
            flash(f"Error al leer archivo: {ex}", 'danger')
            return redirect(url_for('upload', tipo=tipo))

        # — RESTRICCIÓN POR ROL: para asistencia solo sus agentes/admin —
        if tipo=='asistencia' and not current_user.is_admin:
            autorizados = {a.usuario for a in current_user.agentes} if current_user.is_supervisor else {current_user.usuario}
            antes = len(df)
            df = df[df['Usuario'].astype(str).isin(autorizados)]
            if len(df)<antes:
                flash(f"{antes-len(df)} fila(s) omitida(s): fuera de tu cartera.", 'warning')

        cols = ['Usuario','Fecha','Tipo','Monto']
        if tipo=='asistencia':
            cols = ['Usuario','Fecha','Estado','SUP','CARTERA']
        falt = set(cols) - set(df.columns)
        if falt:
            flash(f"Faltan columnas: {', '.join(falt)}", 'warning')
            return redirect(url_for('upload', tipo=tipo))

        for i,row in df.iterrows():
            fila = i+2
            try:
                usr   = str(row['Usuario']).strip()
                fecha = pd.to_datetime(row['Fecha']).date()
                # … lógica idéntica a la que ya tenías …
                if tipo=='asistencia':
                    est = str(row['Estado']).strip().upper()
                    sup = str(row['SUP']).strip()
                    car = str(row['CARTERA']).strip()
                    ex = Attendance.query.filter_by(usuario=usr, fecha=fecha).first()
                    if ex:
                        detalles.append(f"Fila {fila}: {usr} {fecha} {ex.estado}→{est}")
                        ex.estado, ex.supervisor, ex.cartera = est, sup, car
                        reemplazados += 1
                    else:
                        db.session.add(Attendance(
                            usuario=usr, fecha=fecha,
                            estado=est, supervisor=sup, cartera=car
                        ))
                        cargados += 1
                elif tipo=='deducciones':
                    monto = float(row['Monto'])
                    db.session.add(Deduction(
                        usuario=usr, fecha=fecha,
                        tipo=str(row['Tipo']).strip(),
                        monto=monto,
                        observacion=str(row.get('Observación','')).strip()
                    ))
                    cargados += 1
                else:
                    monto = float(row['Monto'])
                    db.session.add(Bonus(
                        usuario=usr, fecha=fecha,
                        tipo=str(row['Tipo']).strip(),
                        monto=monto,
                        observacion=str(row.get('Observación','')).strip()
                    ))
                    cargados += 1
            except Exception as ex:
                errores.append(f"Fila {fila}: {ex}")

        db.session.commit()
        if cargados:     flash(f"Se añadieron {cargados} registros de «{tipo}».", 'success')
        if reemplazados: flash(f"Se reemplazaron {reemplazados} registros de asistencia.", 'info')
        if errores:
            flash("Se omitieron filas con errores:", 'warning')
            for e in errores: flash(e,'warning')
        return redirect(url_for('upload', tipo=tipo))

    return render_template('upload.html', tipo=tipo, plantilla=plantilla)


# —————————————————————————————————————————————————————————
# Listado y filtro de Asistencia + resumen matricial
# —————————————————————————————————————————————————————————
@app.route('/attendance')
@login_required
def list_attendance():
    q = Attendance.query
    if not current_user.is_admin:
        if current_user.is_supervisor:
            allowed = [a.usuario for a in current_user.agentes]
            q = q.filter(Attendance.usuario.in_(allowed))
        else:
            q = q.filter_by(usuario=current_user.usuario)
    records = q.order_by(Attendance.fecha.desc()).all()
    return render_template('attendance.html', records=records)


@app.route('/attendance/filter', methods=['GET'])
@login_required
def filter_attendance():
    usuarios    = [e.usuario for e in Employee.query.all()]
    supervisors = [s[0] for s in db.session.query(Attendance.supervisor).distinct() if s[0]]
    carteras    = [c[0] for c in db.session.query(Attendance.cartera).distinct() if c[0]]

    u_sel = request.args.get('usuario','')
    s_sel = request.args.get('supervisor','')
    c_sel = request.args.get('cartera','')
    d1    = request.args.get('start_date','')
    d2    = request.args.get('end_date','')

    q = Attendance.query
    if not current_user.is_admin:
        if current_user.is_supervisor:
            allowed = [a.usuario for a in current_user.agentes]
            q = q.filter(Attendance.usuario.in_(allowed))
        else:
            q = q.filter_by(usuario=current_user.usuario)

    if u_sel: q = q.filter_by(usuario=u_sel)
    if s_sel: q = q.filter_by(supervisor=s_sel)
    if c_sel: q = q.filter_by(cartera=c_sel)
    if d1:    q = q.filter(Attendance.fecha >= datetime.strptime(d1,'%Y-%m-%d').date())
    if d2:    q = q.filter(Attendance.fecha <= datetime.strptime(d2,'%Y-%m-%d').date())

    registros = q.order_by(Attendance.fecha.desc()).all()

    if d1 and d2:
        start = datetime.strptime(d1,'%Y-%m-%d').date()
        end   = datetime.strptime(d2,'%Y-%m-%d').date()
        date_list = [start+timedelta(days=i) for i in range((end-start).days+1)]
    else:
        date_list = []

    matrix  = defaultdict(lambda: defaultdict(lambda: None))
    summary = defaultdict(Counter)
    for r in registros:
        summary[r.usuario][r.estado] += 1
        matrix[r.usuario][r.fecha]    = r.estado

    observations = []

    return render_template('attendance_filter.html',
        registros     = registros,
        usuarios      = usuarios,
        supervisors   = supervisors,
        carteras      = carteras,
        usuario_sel   = u_sel,
        supervisor_sel= s_sel,
        cartera_sel   = c_sel,
        start_date    = d1,
        end_date      = d2,
        date_list     = date_list,
        matrix        = matrix,
        summary       = summary,
        observations  = observations
    )


@app.route('/attendance/<int:id>/edit', methods=['GET','POST'])
@login_required
def edit_attendance(id):
    r = Attendance.query.get_or_404(id)
    if request.method=='POST':
        nuevo = request.form['estado'].strip().upper()
        if nuevo not in ('A','V','MG','F','D'):
            flash("Estado inválido", "warning")
        else:
            r.estado = nuevo
            r.supervisor = request.form['supervisor'].strip()
            r.cartera = request.form['cartera'].strip()
            db.session.commit()
            flash('Asistencia actualizada.', 'success')
        return redirect(request.referrer or url_for('filter_attendance'))
    return render_template('attendance_edit.html', r=r)


@app.route('/attendance/<int:id>/delete', methods=['POST'])
@login_required
def delete_attendance(id):
    r = Attendance.query.get_or_404(id)
    db.session.delete(r)
    db.session.commit()
    flash('Registro eliminado.', 'warning')
    return redirect(request.referrer or url_for('filter_attendance'))


# —————————————————————————————————————————————————————————
# Listado y filtro de Deducciones y Bonos (idéntico patrón)
# —————————————————————————————————————————————————————————
@app.route('/deductions')
@login_required
def list_deductions():
    records = Deduction.query.order_by(Deduction.fecha.desc()).all()
    return render_template('deductions.html', records=records)

@app.route('/deductions/filter', methods=['GET'])
@login_required
def filter_deductions():
    usuarios = [e.usuario for e in Employee.query.all()]
    u_sel    = request.args.get('usuario','')
    d1       = request.args.get('start_date','')
    d2       = request.args.get('end_date','')
    q = Deduction.query
    if u_sel: q = q.filter_by(usuario=u_sel)
    if d1:    q = q.filter(Deduction.fecha >= datetime.strptime(d1,'%Y-%m-%d').date())
    if d2:    q = q.filter(Deduction.fecha <= datetime.strptime(d2,'%Y-%m-%d').date())
    regs = q.order_by(Deduction.fecha.desc()).all()
    return render_template('deductions_filter.html',
                           registros=regs,
                           usuarios=usuarios,
                           usuario_sel=u_sel,
                           start_date=d1, end_date=d2)

@app.route('/bonuses')
@login_required
def list_bonuses():
    records = Bonus.query.order_by(Bonus.fecha.desc()).all()
    return render_template('bonuses.html', records=records)

@app.route('/bonuses/filter', methods=['GET'])
@login_required
def filter_bonuses():
    usuarios = [e.usuario for e in Employee.query.all()]
    u_sel    = request.args.get('usuario','')
    d1       = request.args.get('start_date','')
    d2       = request.args.get('end_date','')
    q = Bonus.query
    if u_sel: q = q.filter_by(usuario=u_sel)
    if d1:    q = q.filter(Bonus.fecha >= datetime.strptime(d1,'%Y-%m-%d').date())
    if d2:    q = q.filter(Bonus.fecha <= datetime.strptime(d2,'%Y-%m-%d').date())
    regs = q.order_by(Bonus.fecha.desc()).all()
    return render_template('bonuses_filter.html',
                           registros=regs,
                           usuarios=usuarios,
                           usuario_sel=u_sel,
                           start_date=d1, end_date=d2)


# —————————————————————————————————————————————————————————
# Generación de Nómina y exportación a Excel
# —————————————————————————————————————————————————————————
@app.route('/payroll/create', methods=['GET','POST'])
@login_required
def create_payroll():
    if request.method=='POST':
        inicio = datetime.strptime(request.form['inicio'],'%Y-%m-%d').date()
        fin    = datetime.strptime(request.form['fin'],   '%Y-%m-%d').date()
        Payroll.query.filter_by(inicio=inicio, fin=fin).delete()
        db.session.commit()
        for emp in Employee.query.all():
            full = Attendance.query.filter_by(usuario=emp.usuario)\
                   .filter(Attendance.fecha.between(inicio,fin))\
                   .filter(Attendance.estado.in_(['A','V'])).count()
            half = Attendance.query.filter_by(usuario=emp.usuario)\
                   .filter(Attendance.fecha.between(inicio,fin), Attendance.estado=='MG').count()
            days = full + half*0.5
            sb   = emp.salario_diario * days
            tb   = sum(b.monto for b in Bonus.query.filter_by(usuario=emp.usuario)\
                       .filter(Bonus.fecha.between(inicio,fin)))
            td   = sum(d.monto for d in Deduction.query.filter_by(usuario=emp.usuario)\
                       .filter(Deduction.fecha.between(inicio,fin)))
            net  = sb + tb - td
            db.session.add(Payroll(
                usuario           = emp.usuario,
                inicio            = inicio,
                fin               = fin,
                sueldo_base       = sb,
                total_bonos       = tb,
                total_deducciones = td,
                neto              = net
            ))
        db.session.commit()
        return redirect(url_for('list_payroll', inicio=inicio, fin=fin))
    return render_template('payroll_form.html')


@app.route('/payroll', methods=['GET'])
@login_required
def list_payroll():
    inicio = request.args.get('inicio','')
    fin    = request.args.get('fin','')
    try:
        i_date = datetime.strptime(inicio,'%Y-%m-%d').date()
        f_date = datetime.strptime(fin,   '%Y-%m-%d').date()
    except:
        flash("Fechas inválidas", "danger")
        return redirect(url_for('dashboard'))

    records = Payroll.query.filter_by(inicio=i_date, fin=f_date).all()
    total_sb  = sum(r.sueldo_base       for r in records)
    total_bonos       = sum(r.total_bonos       for r in records)
    total_deducciones = sum(r.total_deducciones for r in records)
    total_neto        = sum(r.neto              for r in records)

    return render_template('payroll_list.html',
        records=records,
        inicio=i_date, fin=f_date,
        total_sb=total_sb,
        total_bonos=total_bonos,
        total_deducciones=total_deducciones,
        total_neto=total_neto
    )


@app.route('/payroll/export', methods=['GET'])
@login_required
def export_payroll_xlsx():
    inicio = request.args.get('inicio','')
    fin    = request.args.get('fin','')
    i_date = datetime.strptime(inicio,'%Y-%m-%d').date()
    f_date = datetime.strptime(fin,   '%Y-%m-%d').date()

    records = Payroll.query.filter_by(inicio=i_date, fin=f_date).all()
    data = [{
        'Usuario':     r.usuario,
        'Nombre':      r.employee.nombre,
        'Sueldo Base': r.sueldo_base,
        'Bonos':       r.total_bonos,
        'Deducciones': r.total_deducciones,
        'Neto':        r.neto
    } for r in records]
    df = pd.DataFrame(data)
    totals = {
        'Usuario':     'Totales',
        'Nombre':      '',
        'Sueldo Base': df['Sueldo Base'].sum(),
        'Bonos':       df['Bonos'].sum(),
        'Deducciones': df['Deducciones'].sum(),
        'Neto':        df['Neto'].sum(),
    }
    df = pd.concat([df, pd.DataFrame([totals])], ignore_index=True)

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Nómina')
    output.seek(0)

    filename = f"nomina_{inicio}_a_{fin}.xlsx"
    return Response(
        output.read(),
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        headers={'Content-Disposition': f'attachment; filename={filename}'}
    )


if __name__ == '__main__':
    with app.app_context():
        db.create_all()
    app.run(debug=True)
