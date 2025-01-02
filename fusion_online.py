from flask import Flask

app = Flask(__name__)

@app.route('/')
def home():
    return "Bienvenue sur ma première page Flask ! 🎉"

if __name__ == '__main__':
    app.run(debug=True)
