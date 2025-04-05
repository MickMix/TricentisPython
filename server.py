from flask import Flask, jsonify, request, render_template
from flask_restful import Resource, Api
from flask_cors import CORS
# from sudoku import SudokuGenerator

app = Flask(__name__)
api = Api(app)
CORS(app) # Enable CORS for all routes

@app.route('/tricentis')
def index(): 
    return render_template('index.html')

class GetSudoku(Resource):
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

api.add_resource(GetSudoku, '/generate_sudoku')

if __name__ == '__main__':
    app.run() 
