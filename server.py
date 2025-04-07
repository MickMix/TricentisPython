import os
from flask import Flask, jsonify, request, render_template
from flask_restful import Resource, Api
from flask_cors import CORS
from werkzeug.utils import secure_filename
from pathlib import Path
from waterfall import PipelineWaterfall
# from sudoku import SudokuGenerator

app = Flask(__name__)
api = Api(app)
CORS(app) # Enable CORS for all routes

UPLOAD_FOLDER = 'uploads'
ALLOWED_EXTENSIONS = {'xls', 'xlsx'}

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

@app.route('/tricentis')
def index(): 
    return render_template('index.html')

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/tricentis/submit_form', methods=['POST'])
def submit_form():
    if request.method == 'POST':
        if 'excel-file' not in request.files:
            return 'No file part'
        
        file = request.files['excel-file']
        
        if file.filename == '':
            return 'No selected file'
        
        
        #     print(f"File '{file_path}' does not exist (using pathlib).")
        
        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file_path_pathlib = Path(file_path)

            if file_path_pathlib.exists():
                print(f"File '{file_path}' exists (using pathlib).")
            else:
                file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))

            
            pipeline = PipelineWaterfall()
            outputData = pipeline.buildWaterFall(filename, app.config['UPLOAD_FOLDER'])
            return jsonify(outputData)

class getOutput(Resource):
    def get(self):
        # Get data from JavaScript (e.g., using query parameters)
        # difficulty = request.args.get('difficulty', type=str)
        # generator = SudokuGenerator()
        # boards = generator.generate_board(difficulty)
        # data = jsonify(boards)
        # return data
        return

    # def post(self):
    #     # Define your logic for POST request
    #     data = request.get_json()
    #     return {'message': 'POST method called', 'data': data}, 201

api.add_resource(getOutput, '/get_output')

if __name__ == '__main__':
    os.makedirs(UPLOAD_FOLDER, exist_ok=True)
    app.run() 
