# system path related
import os

# to open or navigate url
import urllib.request

# flask related headers
from flask import Flask, flash, request, redirect, render_template, send_from_directory

# to get secured filename
from werkzeug.utils import secure_filename

# file_converter python script imported
import file_converter


UPLOAD_FOLDER = 'uploaded_files' 			  # folder location where files will be uploaded

# flask app configs
app = Flask(__name__)						  # app declared
app.secret_key = "secret key"				  # secret key defined
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER   # path to folder provided

# to open or navigate to html page
@app.route('/')
def upload_form():
	"""
		this method is used to open
		or navigate to the html 
		file
	"""
	return render_template('upload.html')

# upload files and generate excel file
@app.route('/', methods=['POST'])
def upload_file():
	"""
		this method performs two tasks i.e. 
		firstly it uploads the files to the 
		folder and then process those files 
		to get the excel file generated.
	"""
	# to upload files
	if request.method == 'POST' and "submit_files" in request.form:
		# checks if the post request has the files part
		if 'files[]' not in request.files:
			flash('No file part')
			return redirect(request.url)
		files = request.files.getlist('files[]')
		for file in files:
			if file and allowed_file(file.filename):
				filename = secure_filename(file.filename)
				file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
		flash('File(s) successfully uploaded in the uploaded_files folder')
		return redirect('/')

	# to process files and generate excel
	if request.method == 'POST' and "process_files" in request.form:
		file_converter.main() 		# main method from the python file to convert
		uploads = os.path.join(app.root_path)
		return send_from_directory(directory=uploads, filename='extracted_text.xlsx', as_attachment=True)
 		
	
if __name__ == "__main__":
	app.run(debug=True)