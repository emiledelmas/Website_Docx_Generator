# Website Docx Generator in Python with Flask

This project generates a docx with the data that an user entered into an html form !

## Getting Started

To make this site up you need to install some libraries which are essential to the website and the generator ! 

### Prerequisites

You need to install Pip3 to install these libraries : 

In Windows :
```
pip install flask
pip install python-docx
pip install wikipedia
```
In Linux :
```
$ pip3 install flask
$ pip3 install python-docx
$ pip3 install wikipedia
```

## How modify it

It's a specific app so you need to understand how the website works for using it for your own usage !

### Template and form

The main page (and the only one) is in templates/index.html, the design is based on this template :
[Contact Form v5](https://github.com/lululinda/weapp/tree/master/Lista%20de%20asistencia/ContactFrom_v5%202)

```html
<div class="wrap-input100 bg1 rs1-input100">
<span class="label-input100">Problématique</span>
<input class="input100" type="text" name="problematique" placeholder="Entrer la problématique">
</div>
```

### Flask and docx generator

The script in python opens a docx document which is in templates-docx/ folder. 
In this specific project, it's a 2 or 3 column table which is filled with the data that the user entered in.

So first the script gets what the user put into the form and we put it into variables :

```python
@app.route('/', methods=['POST']) # When the user click to the "submit" button
def my_form_post():
# First we put into variables all the data that the user entered
    titre = request.form['titre']
    author = request.form['auteur']
    number = request.form['number']
    problematique = request.form['problematique']
    nbrePartie = request.form['nbrePartie']
    I = request.form['I']
    I1 = request.form['I1']
    I2 = request.form['I2']
    I3 = request.form['I3']
    II = request.form['II']
    II1 = request.form['II1']
    II2 = request.form['II2']
    II3 = request.form['II3']
    III = request.form['III']
    III1 = request.form['III1']
    III2 = request.form['III2']
    III3 = request.form['III3']
```

Then they are this really useful function which is delete_paragraph :

When we created tables to the docx template (manually) it also created an empty paragraph
and when we add text into the docx file using python-docx library the empty first paragraph is still there and
the text that the user entered in is not on the first line but in the second so first you have to delete empty paragraph
using this function.

```python
def delete_paragraph(paragraph):
    p = paragraph._element
    p.getparent().remove(p)
    p._p = p._element = None
```

Then we can add paragraph to our docx using that function and python-docx library

```python
getTitle = document.tables[0]
        delete_paragraph(getTitle.rows[0].cells[0].paragraphs[-1])
        getTitle.rows[0].cells[0].add_paragraph(titre, style="GrandTitre")
```
The style named "GrandTitre" was already created in the template file, if you want to put your own style you can do it
directly into the docx template file.

## Test it !

You can test the website in this URL : [Docx Generator](https://veksor.pythonanywhere.com/)

## Built With

* [Flask](http://flask.pocoo.org/) - The light python web framework that makes this website that simple and that powerful
* [python-docx](https://python-docx.readthedocs.io/en/latest/) - Python libray that can create, modify and generate docx file easily
* [Wikipedia python library](https://pypi.org/project/wikipedia/) - Used to research the biography of writers
* [Contact Form v5](https://github.com/lululinda/weapp/tree/master/Lista%20de%20asistencia/ContactFrom_v5%202) - Used to create the Html template file

## Authors

* **Emile Delmas** - *Students* - [emiledelmas](https://github.com/emiledelmas)

## License

This project is licensed under the MIT License - see the [LICENSE.md](LICENSE.md) file for details

