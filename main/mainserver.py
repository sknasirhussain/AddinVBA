from flask import Flask
from calculations import app1
from stocks import app2

combined_app = Flask(__name__)

combined_app.register_blueprint(app1)
combined_app.register_blueprint(app2)

if __name__ == "__main__":
    combined_app.run(debug = True)