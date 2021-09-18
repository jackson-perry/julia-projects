from flask import Flask, render_template, request
from werkzeug import secure_filename
app = Flask(__name__)

@app.route('/upload')
def upload_file():
   return render_template('upload.html')
	
@app.route('/uploader', methods = ['GET', 'POST'])
def upload_file():
   if request.method == 'POST':
      f = request.files['file']
      if file:
         file.save('/var/www/PythonProgramming/', DOCXupload.docx)
      return send_file('/var/www/PythonProgramming/index.xlsx', attachment_filename='ohhey.pdf')
		
if __name__ == '__main__':
   app.run(debug = True)
