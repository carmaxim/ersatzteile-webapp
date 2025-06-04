from flask import Flask, render_template_string, request, redirect, send_file
import pandas as pd
import os
from datetime import datetime

app = Flask(__name__)

DATA_FILE = "data.xlsx"

def init_excel():
    if not os.path.exists(DATA_FILE):
        lager_df = pd.DataFrame(columns=["Bezeichnung", "Bestand", "Besitzer"])
        buchungen_df = pd.DataFrame(columns=["Datum", "Bezeichnung", "Menge", "Typ", "Besitzer", "Kommentar"])
        with pd.ExcelWriter(DATA_FILE) as writer:
            lager_df.to_excel(writer, sheet_name="Lager", index=False)
            buchungen_df.to_excel(writer, sheet_name="Historie", index=False)

init_excel()

def load_data():
    xls = pd.ExcelFile(DATA_FILE)
    return xls.parse("Lager"), xls.parse("Historie")

def save_data(lager_df, historie_df):
    with pd.ExcelWriter(DATA_FILE) as writer:
        lager_df.to_excel(writer, sheet_name="Lager", index=False)
        historie_df.to_excel(writer, sheet_name="Historie", index=False)

HTML = """
<!doctype html>
<title>Materialverwaltung</title>
<h1>Ersatzteile Lager</h1>
<table border=1>
  <tr><th>Bezeichnung</th><th>Bestand</th><th>Besitzer</th></tr>
  {% for row in lager.itertuples() %}
  <tr><td>{{ row.Bezeichnung }}</td><td>{{ row.Bestand }}</td><td>{{ row.Besitzer }}</td></tr>
  {% endfor %}
</table>

<h2>Buchung</h2>
<form method=post action="/buchen">
  Bezeichnung: <input name=bezeichnung required><br>
  Menge: <input name=menge type=number required><br>
  Typ: <select name=typ><option value="Ein">Ein</option><option value="Aus">Aus</option></select><br>
  Besitzer: <input name=besitzer required><br>
  Kommentar: <input name=kommentar><br>
  <button type=submit>Buchen</button>
</form>

<h2>Historie</h2>
<table border=1>
  <tr><th>Datum</th><th>Bezeichnung</th><th>Menge</th><th>Typ</th><th>Besitzer</th><th>Kommentar</th></tr>
  {% for row in historie.itertuples() %}
  <tr>
    <td>{{ row.Datum }}</td><td>{{ row.Bezeichnung }}</td><td>{{ row.Menge }}</td>
    <td>{{ row.Typ }}</td><td>{{ row.Besitzer }}</td><td>{{ row.Kommentar }}</td>
  </tr>
  {% endfor %}
</table>

<p><a href="/export">Excel Export</a></p>
"""

@app.route("/")
def index():
    lager, historie = load_data()
    return render_template_string(HTML, lager=lager, historie=historie)

@app.route("/buchen", methods=["POST"])
def buchen():
    bezeichnung = request.form["bezeichnung"]
    menge = int(request.form["menge"])
    typ = request.form["typ"]
    besitzer = request.form["besitzer"]
    kommentar = request.form.get("kommentar", "")

    lager, historie = load_data()
    if bezeichnung not in lager.Bezeichnung.values:
        lager = pd.concat([lager, pd.DataFrame([[bezeichnung, 0, besitzer]], columns=lager.columns)])

    idx = lager[lager.Bezeichnung == bezeichnung].index[0]
    bestand = lager.loc[idx, "Bestand"]
    if typ == "Ein":
        lager.loc[idx, "Bestand"] = bestand + menge
    else:
        lager.loc[idx, "Bestand"] = max(0, bestand - menge)

    historie = pd.concat([historie, pd.DataFrame([[datetime.now(), bezeichnung, menge, typ, besitzer, kommentar]], columns=historie.columns)])

    save_data(lager, historie)
    return redirect("/")

@app.route("/export")
def export():
    return send_file(DATA_FILE, as_attachment=True)

if __name__ == "__main__":
    import os
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
