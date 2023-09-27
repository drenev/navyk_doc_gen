from flask import Flask, request, send_file
import time
from generator import create_dict, create_new_file

app = Flask(__name__, static_folder='static')


@app.route('/')
def index():
    return open('index.html').read()


@app.route('/submit', methods=['POST'])
def submit():
    form_data = {}
    form_data['timestamp'] = time.time() * 1000000
    for field_name, field_value in request.form.items():
        form_data[field_name] = field_value
    print(form_data)
    if form_data:
        create_new_file(create_dict(form_data))

    doc_name = str(form_data['parent_name']) + '__' + str(form_data['timestamp']) + '.docx'
    print(doc_name)

    return send_file(doc_name, as_attachment=True)


if __name__ == '__main__':
    app.run()
