import os
from flask import Flask, render_template, request, redirect, url_for, flash, send_file, jsonify
from flask_sqlalchemy import SQLAlchemy
import pandas as pd
from datetime import datetime, date, timedelta, time
from io import BytesIO

# WICHTIG: Diese Zeile muss VOR dem Import von pyplot stehen!
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt

from fpdf import FPDF
from io import BytesIO

# --- App Konfiguration ---
app = Flask(__name__)
# Ein geheimer Schlüssel ist für Flash-Nachrichten und Sessions erforderlich
app.config['SECRET_KEY'] = 'dein_super_geheimer_schluessel_12345'
basedir = os.path.abspath(os.path.dirname(__file__))
# Konfiguration für die SQLite-Datenbank
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///' + os.path.join(basedir, 'zeiterfassung.db')
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False

db = SQLAlchemy(app)

# --- Datenbankmodelle ---
class TimeEntry(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    date = db.Column(db.Date, nullable=False, default=date.today)
    start_time = db.Column(db.Time, nullable=False)
    end_time = db.Column(db.Time, nullable=False)
    category = db.Column(db.String(50), nullable=False)
    project = db.Column(db.String(100), nullable=False)
    info_text = db.Column(db.String(300), nullable=True)

    @property
    def duration(self):
        start_dt = datetime.combine(self.date, self.start_time)
        end_dt = datetime.combine(self.date, self.end_time)
        if end_dt < start_dt:
            end_dt += timedelta(days=1)
        return end_dt - start_dt

    @property
    def duration_str(self):
        total_seconds = self.duration.total_seconds()
        hours, remainder = divmod(total_seconds, 3600)
        minutes, _ = divmod(remainder, 60)
        return f"{int(hours):02}:{int(minutes):02}"
        
class QuestionAnswer(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    question = db.Column(db.String(500), nullable=False, unique=True)
    options = db.Column(db.String(500), nullable=False)
    answer = db.Column(db.String(50), nullable=True)

# --- Kontext-Prozessor ---
@app.context_processor
def inject_now():
    return {'now': datetime.utcnow()}

# --- Routen / Unterseiten ---
@app.route('/')
def index():
    return render_template('index.html')

@app.route('/begriffsfinder', methods=['GET', 'POST'])
def begriffsfinder():
    search_term = ""
    result = ""
    if request.method == 'POST':
        search_term = request.form.get('search_term', '').strip()
        if not search_term:
            flash("Bitte geben Sie einen Suchbegriff ein.", "error")
        else:
            try:
                df = pd.read_excel('Daten.xlsx', header=None)
                search_area = df.iloc[13:]
                match = search_area[search_area[3].astype(str).str.lower() == search_term.lower()]
                if not match.empty:
                    explanation = match.iloc[0, 7]
                    result = str(explanation) if pd.notna(explanation) else "Keine Erklärung für diesen Begriff vorhanden."
                else:
                    result = "Begriff nicht gefunden."
            except FileNotFoundError:
                result = "Fehler: Die Datei 'begriffe.xlsx' wurde nicht im Hauptverzeichnis gefunden."
            except Exception as e:
                result = f"Ein unerwarteter Fehler ist aufgetreten: {e}"
    return render_template('begriffsfinder.html', search_term=search_term, result=result)

@app.route('/autocomplete_begriffe')
def autocomplete_begriffe():
    query = request.args.get('q', '').lower()
    suggestions = []
    if query:
        try:
            df = pd.read_excel('Daten.xlsx', header=None)
            matching_terms = df[df[3].astype(str).str.lower().str.startswith(query)].iloc[:, 3].unique()
            suggestions = matching_terms.tolist()
            suggestions = suggestions[:10]
        except FileNotFoundError:
            suggestions = ["Fehler: 'begriffe.xlsx' nicht gefunden."]
        except Exception as e:
            suggestions = [f"Ein Fehler ist aufgetreten: {e}"]
    return jsonify(suggestions)

@app.route('/dokumentation', methods=['GET', 'POST'])
def dokumentation():
    if request.method == 'POST':
        try:
            date_obj = datetime.strptime(request.form.get('date'), '%Y-%m-%d').date()
            start_time_obj = datetime.strptime(request.form.get('start_time'), '%H:%M').time()
            end_time_obj = datetime.strptime(request.form.get('end_time'), '%H:%M').time()
            category = request.form.get('category')
            project = request.form.get('project')
            info_text = request.form.get('info_text')
            new_entry = TimeEntry(
                date=date_obj,
                start_time=start_time_obj,
                end_time=end_time_obj,
                category=category,
                project=project,
                info_text=info_text
            )
            db.session.add(new_entry)
            db.session.commit()
            flash('Eintrag erfolgreich gespeichert!', 'success')
        except Exception as e:
            flash(f'Fehler beim Speichern: {e}', 'error')
        return redirect(url_for('dokumentation'))

    entries = TimeEntry.query.order_by(TimeEntry.date.desc(), TimeEntry.start_time.desc()).all()
    generate_category_chart(entries)
    return render_template('dokumentation.html', entries=entries)

@app.route('/delete/<int:entry_id>', methods=['POST'])
def delete_entry(entry_id):
    entry_to_delete = TimeEntry.query.get_or_404(entry_id)
    try:
        db.session.delete(entry_to_delete)
        db.session.commit()
        flash('Eintrag wurde gelöscht.', 'success')
    except Exception as e:
        flash(f'Fehler beim Löschen des Eintrags: {e}', 'error')
    return redirect(url_for('dokumentation'))

@app.route('/generate_pdf')
def generate_pdf():
    period = request.args.get('period', 'month')
    report_date_str = request.args.get('report_date', date.today().strftime('%Y-%m-%d'))
    report_date = datetime.strptime(report_date_str, '%Y-%m-%d').date()
    query = TimeEntry.query
    title = "Zeiterfassungsbericht"
    if period == 'day':
        start_date = report_date
        end_date = report_date
        query = query.filter(TimeEntry.date == report_date)
        title += f" für den {start_date.strftime('%d.%m.%Y')}"
    elif period == 'week':
        start_date = report_date - timedelta(days=report_date.weekday())
        end_date = start_date + timedelta(days=6)
        query = query.filter(TimeEntry.date.between(start_date, end_date))
        title += f" für KW{start_date.isocalendar()[1]} ({start_date.strftime('%d.%m.')} - {end_date.strftime('%d.%m.%Y')})"
    elif period == 'month':
        start_date = report_date.replace(day=1)
        next_month = (start_date.replace(day=28) + timedelta(days=4)).replace(day=1)
        end_date = next_month - timedelta(days=1)
        query = query.filter(TimeEntry.date.between(start_date, end_date))
        title += f" für {start_date.strftime('%B %Y')}"
        
    entries = query.order_by(TimeEntry.date.asc(), TimeEntry.start_time.asc()).all()
    if not entries:
        flash(f'Keine Daten für den ausgewählten Zeitraum ({period}) gefunden.', 'info')
        return redirect(url_for('dokumentation'))
    pdf = PDF(orientation='L', unit='mm', format='A4')
    pdf.set_title_text(title)
    pdf.add_page()
    pdf.create_table(entries)
    pdf_output = pdf.output(dest='S')
    return send_file(
        BytesIO(pdf_output),
        as_attachment=True,
        download_name=f'Zeiterfassung_{period}_{report_date.strftime("%Y-%m-%d")}.pdf',
        mimetype='application/pdf'
    )

@app.route('/fragen', methods=['GET', 'POST'])
def fragen():
    if request.method == 'POST':
        if 'new_question' in request.form:
            new_q_text = request.form.get('new_question')
            options_text = request.form.get('options')
            if not new_q_text or not options_text:
                flash("Bitte geben Sie sowohl eine Frage als auch Antwortoptionen ein.", "error")
                return redirect(url_for('fragen'))
            try:
                existing_q = QuestionAnswer.query.filter_by(question=new_q_text).first()
                if existing_q:
                    flash("Diese Frage existiert bereits!", "error")
                else:
                    new_entry = QuestionAnswer(question=new_q_text, options=options_text, answer=None)
                    db.session.add(new_entry)
                    db.session.commit()
                    flash('Frage erfolgreich erstellt!', 'success')
            except Exception as e:
                flash(f'Fehler beim Speichern der Frage: {e}', 'error')
            return redirect(url_for('fragen'))
        
        for key, value in request.form.items():
            if key.startswith('answer_'):
                q_id = key.split('_')[1]
                entry = QuestionAnswer.query.get(q_id)
                if entry:
                    entry.answer = value
        db.session.commit()
        flash('Antworten erfolgreich gespeichert!', 'success')
        return redirect(url_for('fragen'))

    fragen_db = QuestionAnswer.query.all()
    return render_template('fragen.html', fragen=fragen_db)

@app.route('/delete_question/<int:question_id>', methods=['POST'])
def delete_question(question_id):
    question_to_delete = QuestionAnswer.query.get_or_404(question_id)
    try:
        db.session.delete(question_to_delete)
        db.session.commit()
        flash('Frage wurde gelöscht.', 'success')
    except Exception as e:
        flash(f'Fehler beim Löschen der Frage: {e}', 'error')
    return redirect(url_for('fragen'))

@app.route('/download_filtered_pdf', methods=['POST'])
def download_filtered_pdf():
    try:
        df_excel = pd.read_excel('Daten.xlsx', sheet_name='Fragestellungen', header=13)
        
        fragen_db = QuestionAnswer.query.all()

        if not fragen_db:
            flash("Es sind keine Fragen in der Datenbank gespeichert.", "error")
            return redirect(url_for('fragen'))

        filters = {q.question.strip(): q.answer.strip() for q in fragen_db if q.answer is not None}

        if not filters:
            flash("Bitte beantworten Sie zuerst alle Fragen.", "error")
            return redirect(url_for('fragen'))

        filtered_df = df_excel.copy()
        
        excel_cols_normalized = {col.strip().replace('?', ''): col for col in df_excel.columns}

        for question, answer in filters.items():
            normalized_question = question.strip().replace('?', '')
            
            if normalized_question in excel_cols_normalized:
                col_name = excel_cols_normalized[normalized_question]
                filtered_df = filtered_df[filtered_df[col_name].astype(str).str.lower().str.strip() == answer.lower()]
            else:
                flash(f"Die Frage '{question}' wurde nicht in der Excel-Datei gefunden.", "error")
                return redirect(url_for('fragen'))

        if filtered_df.empty:
            flash("Keine Einträge entsprechen den ausgewählten Kriterien.", "info")
            return redirect(url_for('fragen'))
        
        solution_columns_df = filtered_df.iloc[:, 26:]
        solution_headers = filtered_df.columns[26:].tolist()
        
        pdf = FPDF(orientation='L', unit='mm', format='A4')
        pdf.add_page()
        pdf.set_font("Arial", 'B', 16)

        pdf.cell(0, 10, txt="Fragestellungen", ln=True, align='C')
        pdf.ln(5)

        for question, answer in filters.items():
            pdf.set_font("Arial", 'B', 10)
            pdf.cell(0, 7, txt=f"Frage: {question}", ln=True)
            pdf.set_font("Arial", '', 10)
            pdf.cell(0, 7, txt=f"Beantwortet mit: {answer}", ln=True)
            pdf.ln(2)
        
        pdf.ln(10)
        
        pdf.set_font("Arial", 'B', 16)
        pdf.cell(0, 10, txt="Gefilterte Lösungen", ln=True, align='C')
        pdf.ln(5)

        if not solution_columns_df.empty:
            num_solutions = len(solution_headers)
            col_width = pdf.w / (num_solutions + 1)
            
            pdf.set_font("Arial", 'B', 8)
            for header in solution_headers:
                pdf.cell(col_width, 10, txt=str(header), border=1, align='C')
            pdf.ln()

            pdf.set_font("Arial", '', 8)
            for index, row in solution_columns_df.iterrows():
                for value in row:
                    text = str(value) if pd.notna(value) else ""
                    pdf.cell(col_width, 10, txt=text, border=1, align='L')
                pdf.ln()
            
        pdf_output = pdf.output(dest='S')

        return send_file(
            BytesIO(pdf_output),
            as_attachment=True,
            download_name='gefilterte_daten.pdf',
            mimetype='application/pdf'
        )

    except FileNotFoundError:
        flash("Fehler: Die Datei 'Daten.xlsx' oder das Arbeitsblatt 'Fragestellungen' wurde nicht gefunden. Bitte stellen Sie sicher, dass sie im Hauptverzeichnis der Anwendung liegt.", "error")
        return redirect(url_for('fragen'))
    except Exception as e:
        flash(f"Ein unerwarteter Fehler ist aufgetreten: {e}", "error")
        return redirect(url_for('fragen'))

# --- Hilfsfunktionen für Grafik & PDF ---
def generate_category_chart(entries):
    plt.style.use('dark_background')
    fig, ax = plt.subplots(figsize=(10, 6))
    fig.patch.set_facecolor('#1a1a2e')
    ax.set_facecolor('#1a1a2e')
    if not entries:
        ax.text(0.5, 0.5, 'Keine Daten für die Grafik vorhanden', ha='center', va='center', color='white', fontsize=12)
        plt.savefig('static/img/category_chart.png', bbox_inches='tight')
        plt.close(fig)
        return
    category_durations = {}
    for entry in entries:
        duration_in_hours = entry.duration.total_seconds() / 3600
        category_durations[entry.category] = category_durations.get(entry.category, 0) + duration_in_hours
    labels = category_durations.keys()
    sizes = category_durations.values()
    wedges, texts, autotexts = ax.pie(sizes, autopct='%1.1f%%', startangle=140, pctdistance=0.85)
    for text in texts + autotexts:
        text.set_color('white')
    ax.axis('equal')
    ax.set_title('Arbeitsstunden nach Kategorie', color='white', fontsize=16, pad=20)
    ax.legend(wedges, labels, title="Kategorien", loc="center left", bbox_to_anchor=(1, 0, 0.5, 1))
    plt.savefig('static/img/category_chart.png', bbox_inches='tight', pad_inches=0.1)
    plt.close(fig)

class PDF(FPDF):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.title_text = "Zeiterfassungsbericht"
    def set_title_text(self, text):
        self.title_text = text
    def header(self):
        self.set_font('Arial', 'B', 15)
        self.cell(0, 10, self.title_text, 0, 1, 'C')
        self.ln(10)
    def footer(self):
        self.set_y(-15)
        self.set_font('Arial', 'I', 8)
        self.cell(0, 10, f'Seite {self.page_no()}', 0, 0, 'C')
    def create_table(self, data):
        self.set_font('Arial', 'B', 10)
        col_widths = [25, 20, 20, 20, 35, 45, 110]
        header = ['Datum', 'Start', 'Ende', 'Dauer', 'Kategorie', 'Projekt', 'Infotext']
        for i, h in enumerate(header):
            self.cell(col_widths[i], 10, h, 1, 0, 'C')
        self.ln()
        self.set_font('Arial', '', 9)
        total_duration = timedelta()
        for entry in data:
            total_duration += entry.duration
            row = [
                entry.date.strftime('%d.%m.%Y'),
                entry.start_time.strftime('%H:%M'),
                entry.end_time.strftime('%H:%M'),
                entry.duration_str,
                entry.category,
                entry.project,
                entry.info_text or ""
            ]
            start_y = self.get_y()
            self.cell(col_widths[0], 10, row[0], 1)
            self.cell(col_widths[1], 10, row[1], 1)
            self.cell(col_widths[2], 10, row[2], 1)
            self.cell(col_widths[3], 10, row[3], 1)
            self.cell(col_widths[4], 10, row[4], 1)
            self.cell(col_widths[5], 10, row[5], 1)
            self.multi_cell(col_widths[6], 10, row[6], 1)
        self.ln(5)
        self.set_font('Arial', 'B', 10)
        total_seconds = total_duration.total_seconds()
        hours, remainder = divmod(total_seconds, 3600)
        minutes, _ = divmod(remainder, 60)
        self.cell(sum(col_widths[:3]), 10, "Gesamtdauer:", 1)
        self.cell(col_widths[3], 10, f"{int(hours):02}:{int(minutes):02}", 1)
        self.cell(sum(col_widths[4:]), 10, "", 1, 1)

if __name__ == '__main__':
    with app.app_context():
        db.create_all()
    app.run(debug=True)