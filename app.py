import os
from flask import Flask, render_template, request, redirect, url_for, flash, send_file
from flask_sqlalchemy import SQLAlchemy
import pandas as pd
from datetime import datetime, time
import matplotlib.pyplot as plt
from fpdf import FPDF
from io import BytesIO

# --- App Konfiguration ---
app = Flask(__name__)
app.config['SECRET_KEY'] = 'dein_sehr_geheimer_schluessel' # Wichtig für Flash-Nachrichten
basedir = os.path.abspath(os.path.dirname(__file__))
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///' + os.path.join(basedir, 'zeiterfassung.db')
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False

db = SQLAlchemy(app)

# --- Datenbankmodell für die Zeiterfassung ---
class TimeEntry(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    date = db.Column(db.Date, nullable=False)
    start_time = db.Column(db.Time, nullable=False)
    end_time = db.Column(db.Time, nullable=False)
    category = db.Column(db.String(50), nullable=False)
    project = db.Column(db.String(100), nullable=False)
    info_text = db.Column(db.String(300), nullable=True)

    @property
    def duration(self):
        # Berechnet die Dauer als timedelta
        dummy_date = datetime(1, 1, 1)
        start_dt = datetime.combine(dummy_date, self.start_time)
        end_dt = datetime.combine(dummy_date, self.end_time)
        return end_dt - start_dt

    @property
    def duration_str(self):
        # Formatiert die Dauer als HH:MM String
        total_seconds = self.duration.total_seconds()
        hours, remainder = divmod(total_seconds, 3600)
        minutes, _ = divmod(remainder, 60)
        return f"{int(hours):02}:{int(minutes):02}"



class PDF(FPDF):
    def header(self):
        self.set_font('Arial', 'B', 12)
        self.cell(0, 10, 'Zeiterfassungsbericht', 0, 1, 'C')

    def footer(self):
        self.set_y(-15)
        self.set_font('Arial', 'I', 8)
        self.cell(0, 10, f'Seite {self.page_no()}', 0, 0, 'C')


# --- Routen / Unterseiten ---

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/begriffsfinder', methods=['GET', 'POST'])
def begriffsfinder():
    search_term = ""
    result = ""
    if request.method == 'POST':
        search_term = request.form.get('search_term')
        try:
            # Lade die Excel-Datei
            df = pd.read_excel('Daten.xlsx', header=None) # header=None, da wir Spalten per Index ansprechen
            
            # Suche ab Zeile 14 (Index 13) in Spalte D (Index 3)
            # iloc[13:] wählt alle Zeilen ab Index 13 aus
            filtered_df = df.iloc[13:]
            
            # Suche nach dem Begriff
            match = filtered_df[filtered_df[3].str.contains(search_term, case=False, na=False)]
            
            if not match.empty:
                # Nimm den ersten Treffer und die Erklärung aus Spalte H (Index 7)
                result = match.iloc[0, 7]
            else:
                result = "Begriff nicht gefunden."
        except FileNotFoundError:
            result = "Fehler: 'begriffe.xlsx' wurde nicht gefunden."
        except Exception as e:
            result = f"Ein Fehler ist aufgetreten: {e}"
            
    return render_template('begriffsfinder.html', search_term=search_term, result=result)

@app.route('/dokumentation', methods=['GET', 'POST'])
def dokumentation():
    if request.method == 'POST':
        # Daten aus dem Formular holen
        date_str = request.form.get('date')
        start_time_str = request.form.get('start_time')
        end_time_str = request.form.get('end_time')
        category = request.form.get('category')
        project = request.form.get('project')
        info_text = request.form.get('info_text')

        # Konvertierung und Validierung
        if not all([date_str, start_time_str, end_time_str, category, project]):
            flash('Bitte alle Pflichtfelder ausfüllen!', 'error')
            return redirect(url_for('dokumentation'))
        
        try:
            date_obj = datetime.strptime(date_str, '%Y-%m-%d').date()
            start_time_obj = datetime.strptime(start_time_str, '%H:%M').time()
            end_time_obj = datetime.strptime(end_time_str, '%H:%M').time()

            # Neuen Eintrag erstellen und speichern
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
        except ValueError:
            flash('Ungültiges Datum- oder Zeitformat.', 'error')

        return redirect(url_for('dokumentation'))

    # GET Request: Daten anzeigen
    entries = TimeEntry.query.order_by(TimeEntry.date.desc(), TimeEntry.start_time.desc()).all()
    generate_category_chart(entries) # Grafik bei jedem Laden neu erstellen
    return render_template('dokumentation.html', entries=entries)


@app.route('/generate_pdf')
def generate_pdf():
    # Hier könnte man noch Filter (Woche, Monat) einbauen
    entries = TimeEntry.query.order_by(TimeEntry.date.desc()).all()

    if not entries:
        flash('Keine Daten für PDF-Export vorhanden.', 'error')
        return redirect(url_for('dokumentation'))

    pdf = PDF()
    pdf.add_page()
    
    # Tabellenkopf
    pdf.set_font('Arial', 'B', 10)
    col_widths = [25, 20, 20, 25, 30, 30, 40]
    header = ['Datum', 'Start', 'Ende', 'Dauer', 'Kategorie', 'Projekt', 'Infotext']
    for i, h in enumerate(header):
        pdf.cell(col_widths[i], 10, h, 1)
    pdf.ln()

    # Tabelleninhalt
    pdf.set_font('Arial', '', 9)
    for entry in entries:
        pdf.cell(col_widths[0], 10, entry.date.strftime('%d.%m.%Y'), 1)
        pdf.cell(col_widths[1], 10, entry.start_time.strftime('%H:%M'), 1)
        pdf.cell(col_widths[2], 10, entry.end_time.strftime('%H:%M'), 1)
        pdf.cell(col_widths[3], 10, entry.duration_str, 1)
        pdf.cell(col_widths[4], 10, entry.category, 1)
        pdf.cell(col_widths[5], 10, entry.project, 1)
        pdf.multi_cell(col_widths[6], 10, entry.info_text, 1) # multi_cell for text wrapping
        pdf.ln(0) # Zurück zum Zeilenanfang

    # PDF im Speicher erstellen
    pdf_bytes = pdf.output(dest='S').encode('latin1')
    
    return send_file(
        BytesIO(pdf_bytes),
        as_attachment=True,
        download_name='zeiterfassung.pdf',
        mimetype='application/pdf'
    )


@app.route('/delete/<int:entry_id>')
def delete_entry(entry_id):
    entry_to_delete = TimeEntry.query.get_or_404(entry_id)
    try:
        db.session.delete(entry_to_delete)
        db.session.commit()
        flash('Eintrag wurde gelöscht.', 'success')
    except:
        flash('Fehler beim Löschen des Eintrags.', 'error')
    return redirect(url_for('dokumentation'))

# --- Hilfsfunktionen für Grafik & PDF ---

def generate_category_chart(entries):
    if not entries:
        # Falls keine Daten da sind, leere Grafik erstellen
        plt.style.use('dark_background')
        fig, ax = plt.subplots(figsize=(8, 5))
        ax.text(0.5, 0.5, 'Keine Daten für die Grafik vorhanden', ha='center', va='center', color='white')
        ax.set_facecolor('#1a1a2e')
        fig.patch.set_facecolor('#1a1a2e')
        plt.savefig('static/img/category_chart.png')
        plt.close()
        return

    # Daten für die Grafik vorbereiten
    category_durations = {}
    for entry in entries:
        duration_in_hours = entry.duration.total_seconds() / 3600
        category_durations[entry.category] = category_durations.get(entry.category, 0) + duration_in_hours
    
    labels = category_durations.keys()
    sizes = category_durations.values()
    
    # Pie-Chart erstellen
    plt.style.use('dark_background')
    fig, ax = plt.subplots(figsize=(8, 5))
    ax.pie(sizes, labels=labels, autopct='%1.1f%%', startangle=90, colors=plt.cm.Paired.colors)
    ax.axis('equal')  # Stellt sicher, dass der Pie-Chart rund ist
    ax.set_title('Arbeitsstunden nach Kategorie', color='white')
    fig.patch.set_facecolor('#1a1a2e') # Hintergrund der gesamten Grafik

    plt.savefig('static/img/category_chart.png', bbox_inches='tight')
    plt.close()

# --- App starten ---
if __name__ == '__main__':
    with app.app_context():
        db.create_all() # Erstellt die Datenbank und Tabelle, falls sie nicht existieren
    app.run(debug=True)





