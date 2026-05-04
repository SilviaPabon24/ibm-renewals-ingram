"""
App de prueba local — Renovaciones IBM S&S
Corre con: python app.py
Luego abre: http://localhost:5000/renewals/ui
"""

from flask import Flask
from ibm_renewals_blueprint import ibm_renewals_bp

app = Flask(__name__)
app.register_blueprint(ibm_renewals_bp, url_prefix='/renewals')

if __name__ == '__main__':
    app.run(debug=True, port=5000)
