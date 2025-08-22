import os
from flask import Flask, render_template, request, redirect, url_for, flash, send_file, jsonify
from flask_sqlalchemy import SQLAlchemy
import pandas as pd
from datetime import datetime, date, timedelta, time

# WICHTIG: Diese Zeile muss VOR dem Import von pyplot stehen!
import matplotlib
matplotlib.use('Agg') # Setzt das Backend, um GUI-Fehler im Webserver zu vermeiden
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

# --- Datenbankmodell für die Zeiterfassung ---
# Definiert die Struktur der 'time_entry' Tabelle in unserer Datenbank
class TimeEntry(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    date = db.Column(db.Date, nullable=False, default=date.today)
    start_time = db.Column(db.Time, nullable=False)
    end_time = db.Column(db.Time, nullable=False)
    category = db.Column(db.String(50), nullable=False)
    project = db.Column(db.String(100), nullable=False)
    info_text = db.Column(db.String(300), nullable=True)

    # Eigenschaft, um die Dauer dynamisch zu berechnen
    @property
    def duration(self):
        # Kombiniert Datum und Zeit, um die Differenz korrekt zu berechnen, auch über Mitternacht
        start_dt = datetime.combine(self.date, self.start_time)
        end_dt = datetime.combine(self.date, self.end_time)
        if end_dt < start_dt: # Falls die Endzeit am nächsten Tag liegt
            end_dt += timedelta(days=1)
        return end_dt - start_dt

    # Eigenschaft, um die Dauer als formatierten String (HH:MM) zurückzugeben
    @property
    def duration_str(self):
        total_seconds = self.duration.total_seconds()
        hours, remainder = divmod(total_seconds, 3600)
        minutes, _ = divmod(remainder, 60)
        return f"{int(hours):02}:{int(minutes):02}"

# --- Kontext-Prozessor, um `now` in allen Templates verfügbar zu machen ---
@app.context_processor
def inject_now():
    return {'now': datetime.utcnow()}


# --- Routen / Unterseiten ---

@app.route('/')
def index():
    """Zeigt die Startseite an."""
    return render_template('index.html')

@app.route('/begriffsfinder', methods=['GET', 'POST'])
def begriffsfinder():
    """Verarbeitet die Suche im Begriffsfinder."""
    search_term = ""
    result = ""
    if request.method == 'POST':
        search_term = request.form.get('search_term', '').strip()
        if not search_term:
            flash("Bitte geben Sie einen Suchbegriff ein.", "error")
        else:
            try:
                # Lade die Excel-Datei. header=None, da wir Spalten per Index ansprechen
                df = pd.read_excel('Daten.xlsx', header=None)

                # Suche ab Zeile 14 (Index 13) in Spalte D (Index 3)
                # .iloc[13:] wählt alle Zeilen ab Index 13 aus
                search_area = df.iloc[13:]

                # Suche nach exakter Übereinstimmung (Groß-/Kleinschreibung ignorieren)
                match = search_area[search_area[3].astype(str).str.lower() == search_term.lower()]

                if not match.empty:
                    # Nimm den ersten Treffer und die Erklärung aus Spalte H (Index 7)
                    explanation = match.iloc[0, 7]
                    result = str(explanation) if pd.notna(explanation) else "Keine Erklärung für diesen Begriff vorhanden."
                else:
                    result = "Begriff nicht gefunden."
            except FileNotFoundError:
                result = "Fehler: Die Datei 'Daten.xlsx' wurde nicht im Hauptverzeichnis gefunden."
            except Exception as e:
                result = f"Ein unerwarteter Fehler ist aufgetreten: {e}"

    return render_template('begriffsfinder.html', search_term=search_term, result=result)

@app.route('/autocomplete_begriffe')
def autocomplete_begriffe():
    """Liefert Suchvorschläge für den Begriffsfinder."""
    query = request.args.get('q', '').lower()
    suggestions = []

    if query:
        try:
            df = pd.read_excel('Daten.xlsx', header=None)
            # Suche in Spalte D (Index 3) nach Begriffen, die mit der Eingabe beginnen
            matching_terms = df[df[3].astype(str).str.lower().str.startswith(query)].iloc[:, 3].unique()
            suggestions = matching_terms.tolist()
            # Beschränke die Anzahl der Vorschläge, um die Performance zu verbessern
            suggestions = suggestions[:10]
        except FileNotFoundError:
            suggestions = ["Fehler: 'Daten.xlsx' nicht gefunden."]
        except Exception as e:
            suggestions = [f"Ein Fehler ist aufgetreten: {e}"]

    return jsonify(suggestions)


@app.route('/dokumentation', methods=['GET', 'POST'])
def dokumentation():
    """Verarbeitet die Eingabe und Anzeige der Zeiterfassung."""
    if request.method == 'POST':
        # Daten aus dem Formular holen
        try:
            date_obj = datetime.strptime(request.form.get('date'), '%Y-%m-%d').date()
            start_time_obj = datetime.strptime(request.form.get('start_time'), '%H:%M').time()
            end_time_obj = datetime.strptime(request.form.get('end_time'), '%H:%M').time()
            category = request.form.get('category')
            project = request.form.get('project')
            info_text = request.form.get('info_text')

            # Neuen Eintrag erstellen und in der Datenbank speichern
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

    # GET Request: Alle Daten für die Anzeige laden
    entries = TimeEntry.query.order_by(TimeEntry.date.desc(), TimeEntry.start_time.desc()).all()
    generate_category_chart(entries) # Grafik bei jedem Laden neu erstellen
    return render_template('dokumentation.html', entries=entries)

@app.route('/delete/<int:entry_id>', methods=['POST'])
def delete_entry(entry_id):
    """Löscht einen Zeiterfassungseintrag."""
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
    """Erstellt und liefert einen PDF-Bericht basierend auf dem gewählten Zeitraum."""
    period = request.args.get('period', 'month') # Standard ist 'month'
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

    pdf = PDF(orientation='L', unit='mm', format='A4') # Querformat für mehr Platz
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

# --- Hilfsfunktionen für Grafik & PDF ---

def generate_category_chart(entries):
    """Erstellt ein Kuchendiagramm der Arbeitsstunden nach Kategorie."""
    plt.style.use('default')
    fig, ax = plt.subplots(figsize=(10, 6))
    fig.patch.set_facecolor('#f0f0f0')
    ax.set_facecolor('#f0f0f0')
    
    if not entries:
        ax.text(0.5, 0.5, 'Keine Daten für die Grafik vorhanden', ha='center', va='center', color='#333333', fontsize=12)
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
        text.set_color('#333333')
    
    ax.axis('equal')
    ax.set_title('Arbeitsstunden nach Kategorie', color='#333333', fontsize=16, pad=20)
    ax.legend(wedges, labels, title="Kategorien", loc="center left", bbox_to_anchor=(1, 0, 0.5, 1))

    plt.savefig('static/img/category_chart.png', bbox_inches='tight', pad_inches=0.1)
    plt.close(fig)

class PDF(FPDF):
    """Eigene PDF-Klasse mit Header und Footer."""
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
        col_widths = [25, 20, 20, 20, 35, 45, 110] # Angepasst für Querformat
        header = ['Datum', 'Start', 'Ende', 'Dauer', 'Kategorie', 'Projekt', 'Infotext']
        
        # Tabellenkopf
        for i, h in enumerate(header):
            self.cell(col_widths[i], 10, h, 1, 0, 'C')
        self.ln()

        # Tabelleninhalt
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
            
            # Speichere die aktuelle Y-Position, um alle Zellen auf einer Höhe zu halten
            start_y = self.get_y()
            
            # Zeichne die Zellen mit fester Höhe, außer der letzten
            self.cell(col_widths[0], 10, row[0], 1)
            self.cell(col_widths[1], 10, row[1], 1)
            self.cell(col_widths[2], 10, row[2], 1)
            self.cell(col_widths[3], 10, row[3], 1)
            self.cell(col_widths[4], 10, row[4], 1)
            self.cell(col_widths[5], 10, row[5], 1)
            
            # Die multi_cell für den Infotext kann die Höhe verändern
            self.multi_cell(col_widths[6], 10, row[6], 1)
        
        # Gesamtdauer am Ende hinzufügen
        self.ln(5) # Etwas Abstand
        self.set_font('Arial', 'B', 10)
        total_seconds = total_duration.total_seconds()
        hours, remainder = divmod(total_seconds, 3600)
        minutes, _ = divmod(remainder, 60)
        self.cell(sum(col_widths[:3]), 10, "Gesamtdauer:", 1)
        self.cell(col_widths[3], 10, f"{int(hours):02}:{int(minutes):02}", 1)
        self.cell(sum(col_widths[4:]), 10, "", 1, 1) # 1 am Ende für Zeilenumbruch


# --- App starten ---
if __name__ == '__main__':
    with app.app_context():
        # Erstellt die Datenbank und Tabelle, falls sie nicht existieren
        db.create_all()
    app.run(debug=True)