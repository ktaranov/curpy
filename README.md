python3 -m venv venv

source venv/Scripts/activate ( Windows )

pip install -r requirements.txt

python3 main.py

На выходе в директории /curpy/ появится итоговый Excel файл, отсортированный и заполненный по всем правилам.
Содержит 2 листа: Без выходных и С выходными и праздниками. Разница в том, что Без выходных - это склеенные исходники, отсортированные по дате.