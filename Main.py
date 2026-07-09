from fastapi import FastAPI, Request, Form
from fastapi.responses import HTMLResponse
import sqlite3

html = """
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>Registro</title>

        <style>
        input, textarea, select {
            width: 300px;
            padding: 5px;
            margin-bottom: 10px;
        }
    </style>
</head>
<body>
    <form action="/enviar" method="post">
    <div>
    <label for="empresa">Empresa</label>
    <input name="empresa" type="text" id="empresa" required>
    <br><br>
    </div>

    <div>
    <label for="vidas">Vidas</label>
    <input name="vidas" type="number" id="vidas" required>
    <br><br>
    </div>

    <div>
    <label for="status">Status</label>
    <select name="status" id="status" required>
        <option value="Agendado">Agendado</option>
        <option value="Realizado">Realizado</option>
    </select>
    <br><br>
    </div>

    <div>
    <label for="obs">Observação</label>
    <textarea name="observacao" id="obs" rows="3"></textarea>
    <br><br>
    </div>

    <div>
    <label for="contato">Contato</label>
    <input name="contato" type="text" id="contato" required>
    <br><br>
    </div>

    <div>
    <label for="telefone">Telefone</label>
    <input name="numero" type="tel" id="numero" required>
    <br><br>
    </div>

    <div>
    <label for="email">Email</label>
    <input name="email" type="email" id="email" required>
    <br><br>
    </div>

    <div>
    <label for="consultora">Consultora</label>
    <input name="consultora" type="text" id="consultora" required>
    <br><br>
    </div>

    <div>
    <label for="analista">Analista</label>
    <input name="analista" type="text" id="analista" required>
    <br><br>
    </div>

    <div>
    <label for="congenere">Congenere</label>
    <input name="congenere" type="text" id="congenere" required>
    <br><br>
    </div>

    <div>
    <label for="hotphone">HotPhone</label>
    <select name="hotphone" id="hotphone" required>
        <option value="Sim">Sim</option>
        <option value="Nao">Não</option>
    </select>
    <br><br>
    </div>

    <div>
    <label for="carteira">Carteira</label>
    <input name="carteira" type="text" id="carteira" required>
    <br><br>
    </div>

    <div>
    <label for="agendamento">Data De Agendamento</label>
    <input name="agendamento" type="date" id="agendamento" required>
    <br><br>
    </div>

    <div>
    <label for="visita">Data De Visita</label>
    <input name="visita" type="date" id="visita" required>
    <br><br>
    </div>

    <div>
    <label for="horario">Horario da visita</label>
    <input name="horario" type="time" id="horario" required>
    <br><br>
    </div>

    <div>
    <button type="submit">Enviar</button>
    </div>
</form>
</body>
</html>
"""

conn = sqlite3.connect("Leads.db")

conn.execute("""CREATE TABLE IF NOT EXISTS agendamentos (
             id INTEGER PRIMARY KEY AUTOINCREMENT,
             data_agendamento TEXT,
             data_visita TEXT,
             horario TEXT,
             empresa TEXT,
             contato TEXT,
             email TEXT,
             vidas INTEGER,
             congenere TEXT,
             hotphone TEXT,
             status TEXT,
             carteira TEXT,
             observacao TEXT,
             criado_em TIMESTAMP DEFAULT CURRENT_TIMESTAMP)""")

app = FastAPI()

@app.get("/")
def main():
    return HTMLResponse(content=html)

@app.post("/enviar")
def receber_dados(nome: str = Form(...)):
    return print("dados recebidos e registrados!")





if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
