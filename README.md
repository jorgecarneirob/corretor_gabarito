Autores: Jorge Carneiro e Wagner Tomé
jorgecarneiroconsultoria@gmail.com

Corretor de gabaritos:
Esta aplicação foi testada com a versão Python 3.11 e nas distribuições Debian 12 e Fedora 41.

Pré-requisitos:
python3
python3-pip
python3-venv
mesa-libGL (Baseados em debian)
libglvnd-glx (baseados em pacotes rpm)
git

Processo de instalação:
1. Clone este repositório:
git clone https://github.com/jorgecarneirob/corretor_gabarito.git

2. Acesse a pasta do repositório:
cd corretor_gabarito

3. Crie um ambiente virtual para instalar as bibliotecas e dependências:
python3 -m venv venv

4. Instale as dependências:
pip install -r requeriments.txt

5. Inicie a aplicação com o gunicorn:
gunicorn --workers 4 --bind 0.0.0.0:8000 app:app

6. Acesse a aplicação via web:
127.0.0.1:8000 ou seuip:8000


