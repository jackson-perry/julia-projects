# -*- coding: utf-8 -*-
"""
Created on Sat Jun 22 20:32:41 2019

@author: jacks
"""
import os
from flask import Flask, render_template
PROJECT_PATH = os.path.realpath(os.path.dirname(__file__))
app = Flask(__name__)

@app.route('/')
def index():
    return render_template("index.html")

    
@app.route('/cakes')
def cakes():
    return 'Yummy cakes!'

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0')