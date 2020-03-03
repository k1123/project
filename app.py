from flask import Flask, render_template

app = Flask(__name__)

from pymongo import MongoClient
client = MongoClient('localhost', 27017)
db = client.dbsparta


@app.route('/')
def home():
    docs = []

    data = list(db.magicdata.find({}, {'_id': 0}))

    for d in data:
        doc = list(d.values())
        docs.append(doc)

    return render_template('page_2.html', docs=docs)

if __name__ == '__main__':
    app.run('0.0.0.0', port=5000, debug=True)

