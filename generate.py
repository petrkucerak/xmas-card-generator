from docx import Document
from docx.shared import Inches
import csv

PATH="data.csv"

def load_csv(file_path):
    data = []

    try:
        with open(file_path, encoding='utf-8') as csv_file:
            csv_reader = csv.DictReader(csv_file)
            for row in csv_reader:
                data.append(row)
    except FileNotFoundError:
        print(f"Error: File not found at {file_path}")
    except Exception as e:
        print(f"An error occurred: {e}")

    return data

# static data for B type (vykani)
head_b = "V Pardubicích dne 10. prosince 2024"
body_b = ["Milí dobrovolníci, dobrodinci a pomocníci,",
"dovolte mi bych vám poděkoval za vaši práci ve farnosti a zapojení se do života našeho společenství.  Jako poděkování přijměte, prosím, pozvání na mši svatou v pátek 17. ledna v 18:00 v kostele sv. Bartoloměje. Bude slavena za vás všechny, kteří jakýmkoliv způsobem aktivně pomáháte ve farnosti. Po bohoslužbě jste zváni na faru, kde u malého občerstvení bude naše společné setkání pokračovat. ",
"Svoji vděčnost bych vám chtěl vyjádřit i věnováním jednoho citátu z Písma. Ač je dopis psán pro všechny stejný, tento citát je pro každého osobní. Není vylosován, ale vybrán. Za každého z vás jsem se alespoň chvíli modlil, promítal jsem si v duchu váš život a přemýšlel jaká věta z Božího slova vyjadřuje vaši službu či životní situaci. Úmyslem citátu bylo vás potěšit, ocenit vaši službu, nebo dodat odvahu do nového roku. Pokud tomu tak není, a při jeho čtení pociťujete něco jiného, tak se omlouvám. K novoročence přikládám i dopis, který každý rok píši svým přátelům a známým.",
"Děkuji vám za vaši službu a Bůh, který vidí i to co je skryté vám vše štědře odplatí!",
"S vděčností Vám žehná ",
"P. Jan Uhlíř"
]
name_b = ", jako můj osobní dar přijměte toto Boží slovo:"

# static data for A type (tykani)
head_a = "V Pardubicích dne 10. prosince 2024"
body_a = ["Milí dobrovolníci, dobrodinci a pomocníci,"
"dovolte mi bych vám poděkoval za vaši práci ve farnosti a zapojení se do života našeho společenství.  Jako poděkování přijměte, prosím, pozvání na mši svatou v pátek 17. ledna v 18:00 v kostele sv. Bartoloměje. Bude slavena za vás všechny, kteří jakýmkoliv způsobem aktivně pomáháte ve farnosti. Po bohoslužbě jste zváni na faru, kde u malého občerstvení bude naše společné setkání pokračovat. ",
"Svoji vděčnost bych vám chtěl vyjádřit i věnováním jednoho citátu z Písma. Ač je dopis psán pro všechny stejný, tento citát je pro každého osobní. Není vylosován, ale vybrán. Za každého z vás jsem se alespoň chvíli modlil, promítal jsem si v duchu váš život a přemýšlel jaká věta z Božího slova vyjadřuje vaši službu či životní situaci. Úmyslem citátu bylo vás potěšit, ocenit vaši službu, nebo dodat odvahu do nového roku. Pokud tomu tak není, a při jeho čtení pociťujete něco jiného, tak se omlouvám. K novoročence přikládám i dopis, který každý rok píši svým přátelům a známým.",
"Děkuji vám za vaši službu a Bůh, který vidí i to co je skryté vám vše štědře odplatí!",
"S vděčností Vám žehná ",
"P. Jan Uhlíř"
]
name_a = ", jako můj osobní dar přijmi toto Boží slovo:"



if __name__ == "__main__":

    data = load_csv(PATH)

    print(data)

    # document = Document()
    # document.add_page_break()
    # document.save('demo.docx')