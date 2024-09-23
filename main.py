import streamlit as st 
import pandas as pd
import sqlite3
from datetime import datetime
import time
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import re

# Configurações do e-mail
EMAIL_HOST = 'smtp.office365.com'
EMAIL_PORT = 587
EMAIL_USER = 'deivison.dias@brisamarshopping.com.br'  # Aqui é o email responsável por enviar as mensagens aos admnistradores
EMAIL_PASS = 'Brisamar#2302@'  #senha do email

# Lista de e-mails dos administradores
ADMIN_EMAILS = ['deivison.dias@brisamarshopping.com.br']

# Definir a senha para acesso à aba "Liberação Operacional"
SENHA_ACESSO = 'Operacional#2024'  # Altere para a senha desejada

# Definir a senha para acesso à tela aba "Devolução de Chaves"
SENHA_ACESSO_DEVOLVER = 'Operacional#2024'

# Função para conectar o banco de dados
def connect_db():
    conn = sqlite3.connect('solicitacoes_chaves.db', timeout=10) # conectar ao banco de dados com um tempo de espera maior
    return conn

# Função para criar a tabela se não existir
def create_table():
    conn = connect_db()
    cursor = conn.cursor()
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS solicitacoes (
        id INTEGER PRIMARY KEY,
        data_solicitacao TEXT,
        data_retirada TEXT,
        data_entrega TEXT,
        setor_solicitante TEXT,
        chave TEXT,
        status TEXT,
        email_solicitante TEXT,
        nome_aprovador TEXT,
        nome_recebedor TEXT,
        atraso INTEGER,
        data_ultimo_email TEXT           
        )                               
    ''')
    conn.commit()
    conn.close()


def is_valid_email(email):
    # Expressão regular para validação básica de e-mail
    email_regex = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
    return re.match(email_regex, email) is not None

# Função para enviar e-mail
def send_email(to_email, subject, body):
    try:
        with smtplib.SMTP(EMAIL_HOST, EMAIL_PORT) as server:
            server.starttls()
            server.login(EMAIL_USER, EMAIL_PASS)

            #Enviar para uma lista de e-mails
            if isinstance(to_email, list):
                for email in to_email:
                    if is_valid_email(email):
                        msg = MIMEMultipart()
                        msg['From'] = EMAIL_USER
                        msg['To'] = email  # Use a variável correta
                        msg['Subject'] = subject
                        msg.attach(MIMEText(body, 'plain'))
                        server.send_message(msg)
                    else:
                        st.warning(f"Endereço de e-mail inválido: {email}")
            elif is_valid_email(to_email):
                msg = MIMEMultipart()
                msg['From'] = EMAIL_USER
                msg['To'] = to_email
                msg['Subject'] = subject
                msg.attach(MIMEText(body, 'plain'))
                server.send_message(msg)
            else:
                st.warning(f"Endereço de e-mail inválido: {to_email}")
    except smtplib.SMTPDataError as e:
        st.error(f"Erro ao enviar e-mail: {e}")

# Função para adicionar nova solicitação
def add_request(data_retirada, data_entrega, setor_solicitante, email, chave):
    conn = connect_db()
    cursor = conn.cursor()
    cursor.execute('''
        INSERT INTO solicitacoes (data_solicitacao, data_retirada, data_entrega, setor_solicitante, chave, status, email_solicitante)
        VALUES (?, ?, ?, ?, ?, ?, ?)
    ''', (datetime.now().strftime("%d-%m-%Y %H:%M:%S"), data_retirada, data_entrega, setor_solicitante, chave, 'Pendente', email))
    conn.commit()
    conn.close()

    # Enviar e-mail para o solicitante
    subject = "Confirmação de Solicitação de Chave"
    body = f"""
    Olá,

    Sua solicitação de chave foi recebida.

    Data da Solicitação: {datetime.now().strftime("%d-%m-%Y %H:%M:%S")}
    Data de Retirada: {data_retirada}
    Data de Entrega: {data_entrega}
    Setor Solicitante: {setor_solicitante}
    Chave: {chave}

    Status: Pendente

    Aguarde a confirmação do status para retirada da chave.

    Atenciosamente,
    SOSC - Sistema Operacional de Solicitação de Chaves
    """
    send_email(email, subject, body)

    # Enviar e-mail para o administrador
    subject_admin = "Nova Solicitação de Chave"
    body_admin = f"""
    Nova solicitação de chave recebida.

    Data da Solicitação: {datetime.now().strftime("%d-%m-%Y %H:%M:%S")}
    Data de Retirada: {data_retirada}
    Data de Entrega: {data_entrega}
    Setor Solicitante: {setor_solicitante}
    Chave: {chave}
    Solicitante: {email}
    """
    send_email(ADMIN_EMAILS, subject_admin, body_admin)

# Função para obter solicitações
def get_requests():
    conn = connect_db()
    df = pd.read_sql_query("SELECT * FROM solicitacoes", conn)
    conn.close()
    return df

# Função para registrar a devolução
def register_return(index, recebedor):
    conn = connect_db()
    cursor = conn.cursor()
    cursor.execute('''
        UPDATE solicitacoes
        SET status = 'Devolvida', nome_recebedor = ?
        WHERE id = ?
    ''', (recebedor, index))
    conn.commit()
    conn.close()

# Função para verificar se a data de retirada ou entrega já passou
def processar_solicitacao(data_retirada, data_entrega):
    hoje = datetime.now().date()
    if data_retirada < hoje or data_entrega < hoje:
        st.warning("A data de retirada ou entrega já passou. Verifique as datas  antes de prosseguir.")
        time.sleep(4)
        return # Impede que a solicitação seja feita
    else:
        add_request()
                   
def check_and_notify_delays():
    hoje = datetime.now().date()
    conn = connect_db()
    cursor = conn.cursor()

    # Obtem a data do último email enviado
    cursor.execute("SELECT data_ultimo_email FROM solicitacoes")
    result = cursor.fetchone()

    if result and isinstance(result[0], str):
        data_ultimo_email = result[0]
    else:
        data_ultimo_email = None

    # Selecionar todas as chaves liberadas
    cursor.execute("SELECT * FROM solicitacoes WHERE status = 'Liberada'")
    rows = cursor.fetchall()
    
    for row in rows:
        index = row[0] # Assume que o ID da da solicitação é primeira coluna
        data_entrega = datetime.strptime(row[3],"%d-%m-%Y").date() # Data da entrega

        if data_entrega < hoje:
                # Calcular os dias de atraso
                atraso_atual = (hoje - data_entrega).days

                # Verificar o atraso atual no banco de dados
                cursor.execute("SELECT atraso, 'data_ultimo_email' FROM solicitacoes WHERE ID = ?", (index,))
                atraso_atual_db, data_ultimo_email_db = cursor.fetchone()

                if atraso_atual_db is None or int(atraso_atual_db) < atraso_atual:
                    # Atualizar a coluna "Atraso"
                    cursor.execute("UPDATE solicitacoes SET Atraso = ? WHERE ID = ?", (atraso_atual, index))

                    # Verificar a data do último email
                    if data_ultimo_email_db and isinstance (data_ultimo_email_db, str):          
                        try:
                            last_email_date = datetime.strptime(data_ultimo_email_db, "%d-%m-%Y").date()
                            if last_email_date < hoje:
                                # Enviar notificação de atraso
                                email = row[7] # Assume que o email do solicitante está na sétima coluna 
                                subject = "Aviso de Atraso na Devolução de Chave"
                                body = f"""
                                Olá,

                                Notamos que a chave que você solicitou deveria ter sido devolvida em {data_entrega.strftime("%d-%m-%Y")}.
                                Atualmente, a devolução está {atraso_atual} dias em atraso.

                                Solicitamos que a devolução da chave seja feita o mais breve possível.

                                Atenciosamente,
                                SOSC - Sistema Operacional de Solicitação de Chaves
                                """
                                send_email(email, subject, body)

                                # Atualizar a data do último e-mail
                                cursor.execute("UPDATE solicitacoes SET 'data_ultimo_email' = ? WHERE ID = ?", (hoje.strftime("%d-%m-%Y"), index))
                        except ValueError:
                            # Caso a data não seja válida:
                            st.warning(f"Data do último e-mail inválida para a solicitação {index}.")
                    else:
                        # Enviar notificação se não houver registro de e-mail
                        email = row[7]  # Assume que o email do solicitante está na sétima coluna
                        subject = "Aviso de Atraso na Devolução de Chave"
                        body = f"""
                        Olá,

                        Notamos que a chave que você solicitou deveria ter sido devolvida em {data_entrega.strftime("%d-%m-%Y")}.
                        Atualmente, a devolução está {atraso_atual} dias em atraso.

                        Solicitamos que a devolução da chave seja feita o mais breve possível.

                        Atenciosamente,
                        SOSC - Sistema Operacional de Solicitação de Chaves
                        """
                        send_email(email, subject, body)

                        # Atualizar a data do último e-mail
                        cursor.execute("UPDATE solicitacoes SET data_ultimo_email = ? WHERE ID = ?", (hoje.strftime("%d-%m-%Y"), index))
    conn.commit()
    conn.close()

# Função para atualizar o status APROVADO
def update_status_aprovado(index, status, nome_aprovador):
    conn = connect_db()
    cursor = conn.cursor()
    cursor.execute('''
        UPDATE solicitacoes
        SET status = ?, nome_aprovador = ?
        WHERE id = ?     
    ''', (status, nome_aprovador, index))

    # Obtém o email do solicitante
    cursor.execute("SELECT email_solicitante FROM solicitacoes WHERE id = ?", (index,))
    result = cursor.fetchone()

    # Verifica se algum resultado foi retornado
    if result is not None:
        email = result[0] # Pega o primeiro item da tupla
    else:
        email = None # ou uma string padrão, se preferir

    conn.commit()
    conn.close()

    subject = "Atualização sobre sua Solicitação de Chave"
    body = f"""
    Olá,

    Sua solicitação de chave foi aprovada.

    Se apresentar na sala Operacional para retirada da chave com o aprovador.

    Aprovador: {nome_aprovador}

    Atenciosamente,
    SOSC - Sistema Operacional de Solicitação de Chaves
    """
    send_email(email, subject, body)

# Função para atualizar o status da solicitação NEGADA
def update_status_negado(index, status, nome_aprovador):
    conn = connect_db()
    cursor = conn.cursor()
    cursor.execute('''
        UPDATE solicitacoes
        SET status = ?, nome_aprovador = ?
        WHERE id = ?
    ''', (status, nome_aprovador, index))

    # Obtém o email do solicitante
    cursor.execute("SELECT email_solicitante FROM solicitacoes WHERE id = ?", (index,))
    result = cursor.fetchone()

    # Verifica se algum resultado foi retornado
    if result is not None:
        email = result[0]
    else:
        email = None # ou uma string padrão, se preferir

    conn.commit()
    conn.close()

    subject = "Atualização sobre sua Solicitação de Chave"
    body = f"""
    Olá,

    Sua solicitação de chave foi Negada.

    Aprovador: {nome_aprovador}

    No momento a chave não se encontra na sala operacional.

    Atenciosamente,
    SOSC - Sistema Operacional de Solicitação de Chaves
    """
    send_email(email, subject, body)

# Função para registrar a devolução da chave
def register_return(index, recebedor):
    conn = connect_db()
    cursor = conn.cursor()
    # Aqui faremos o cálculo para calcular a coluna atraso
    cursor.execute("SELECT data_entrega FROM solicitacoes WHERE id = ?", (index,))
    data_entrega = cursor.fetchone() # Aqui ele retorna a data da entrega

    if data_entrega and data_entrega[0]: # Verifica se a data da entrega não é None
        data_entrega = data_entrega[0]
        
        if isinstance(data_entrega, str):
            # Se for uma string, converta para datetime, ajustar o formato conforme o necessário
            data_entrega = datetime.strptime(data_entrega, '%d-%m-%Y')
        # Lógica para calcular o atraso em dias
        atraso = (datetime.now() - data_entrega).days
        atraso = atraso if atraso > 0 else 0 # Se não houver atraso insere zero
    else:
        atraso = 0 # Se a data de entrega for None insira o zero
    cursor.execute('''
        UPDATE solicitacoes
        SET status = 'Devolvida', nome_recebedor = ?, atraso = ?
        WHERE id = ?
    ''', (recebedor, atraso, index))
    conn.commit()
    conn.close()

    st.success(f"Chave devolvida registrada com sucesso. Nome do recebedor: {recebedor}")


# Tela de Login liberação das chaves
def login_screen_liberar_chaves():
    st.title("Tela de Login")
    senha = st.text_input("Digite a senha:", type="password", key="senha_login_liberar_chaves")
    
    if st.button("Entrar", key="entrar_liberar"):
        if senha == SENHA_ACESSO:
            st.session_state["logged_in"] = True
            st.success("Login realizado com sucesso!")
            time.sleep(2)
            st.rerun()
        else:
            st.error("Senha incorreta. Tente novamente.")

# Tela de Login devolução das chaves
def login_screen_devolver_chaves():
    st.title("Tela de Login")
    senha = st.text_input("Digite a senha:", type="password", key="senha_login_devolver_chaves")
    
    if st.button("Entrar", key="entrar_devolver"):
        if senha == SENHA_ACESSO_DEVOLVER:
            st.session_state["logged_in"] = True
            st.success("Login realizado com sucesso!")
            time.sleep(2)
            st.rerun()
        else:
            st.error("Senha incorreta. Tente novamente.")



# Tela Principal
def main():
    create_table() #Cria a tabela no banco de dados se não existir

    # Verificar e notificar sobre atrasos
    check_and_notify_delays()
    
    col1, col2, col3 = st.columns([2, 5, 2])
    with col2:
        image_path = "Operacional.png"
        st.image(image_path, use_column_width=True)
        

    
    # Inicializar o estado da sessão
    if "logged_in" not in st.session_state:
        st.session_state["logged_in"] = False

    st.title("Sistema Operacional de Solicitação de Chaves")
    

    # Caminho do arquivo Excel
    excel_file = "solicitacoes_chaves.xlsx"

    # Inicialização das chaves do st.session_state
    if 'email' not in st.session_state:
        st.session_state['email'] = ''
    if 'setor' not in st.session_state:
        st.session_state['setor'] = ''
    if 'data_retirada' not in st.session_state:
        st.session_state['data_retirada'] = None
    if 'data_entrega' not in st.session_state:
        st.session_state['data_entrega'] = None
    if 'chave' not in st.session_state:
        st.session_state['chave'] = ''
    if 'reset' not in st.session_state:
        st.session_state['reset'] = False
    
    # Opções para seleção
    setores_opcoes = ["Operacional", "Administração", "Comercial", "Marketing"]  # Setores da unidade

    # Abas para o Solicitação de Chaves , Liberação Operacional, Devolução Chaves
    tab1, tab2, tab3 = st.tabs(["Solicitação de Chaves", "Liberação Operacional", "Devolução de Chaves"])

    # Garantir que o valor do setor está na lista de opções
    if st.session_state['setor'] not in setores_opcoes:
        st.session_state['setor'] = setores_opcoes[0]


    with tab1:
        st.header("Solicitar Chave")

        if st.session_state['reset']:
            st.session_state['email'] = ''
            st.session_state['setor'] = setores_opcoes[0] #Definir um valor padrão
            st.session_state['data_retirada'] = None
            st.session_state['data_entrega'] = None
            st.session_state['chave'] = ''
            st.session_state['reset'] = False

        # Garantir que o valor do setor está na lista de opções

        setores_opcoes = ["Operacional", "Administração", "Comercial", "Marketing"]  # Exemplo de lista de setores
        setor_default = "Operacional"  # Defina o valor padrão que deseja
        setor_default = setores_opcoes[0]  # Valor padrão
        if st.session_state['setor'] in setores_opcoes:
            setor_default = st.session_state['setor']

        # Verificar se o setor_default está na lista de opções
        if setor_default not in setores_opcoes:
            setor_default = setores_opcoes[0]

        email = st.text_input("Digite seu email", st.session_state['email'], key='email')
        setor = st.selectbox(
            "Selecione o seu setor",
            setores_opcoes,
            index=setores_opcoes.index(st.session_state['setor']),
            key='setor'
        )
        data_retirada = st.date_input("Digite a data da retirada da chave", st.session_state['data_retirada'], key='data_retirada')
        data_entrega = st.date_input("Digite a data da entrega da chave", st.session_state['data_entrega'], key='data_entrega')
        chave = st.text_input("Digite a chave que deseja solicitar", st.session_state['chave'], key='chave')

        if st.button("Solicitar Chave"):
            hoje = datetime.now().date()
            if data_retirada < hoje or data_entrega < hoje:
                processar_solicitacao(data_retirada, data_entrega)
            elif email and setor and data_retirada and data_entrega and chave:
                # Converter as datas para strings no formato correto
                data_retirada_str = data_retirada.strftime("%d-%m-%Y")
                data_entrega_str = data_entrega.strftime("%d-%m-%Y")    

                add_request(data_retirada_str, data_entrega_str, setor, email, chave)
                st.success("Solicitação enviada com sucesso!")
                

                # Ativar o reset dos campos
                st.session_state['reset'] = True

                time.sleep(2)

                st.rerun()

            else:
                st.error("Por favor, preencha todos os campos.")

    with tab2:
        
        if not st.session_state.get("logged_in", False):
            login_screen_liberar_chaves()
        else:
            st.header("Liberação")

            # Obter solicitações pendentes
            pending_requests = get_requests()
            pending_requests = pending_requests[pending_requests['status'] == 'Pendente']

            if not pending_requests.empty:
                for i, request in pending_requests.iterrows():
                    st.write(f"Solicitação: {request['id']}")
                    st.write(f"Data da Solicitação: {request['data_solicitacao']}")
                    st.write(f"Data de Retirada: {request['data_retirada']}")
                    st.write(f"Data de Entrega: {request['data_entrega']}")
                    st.write(f"Setor: {request['setor_solicitante']}")
                    st.write(f"Chave: {request['chave']}")
                    st.write(f"Solicitante: {request['email_solicitante']}")
                        

                    # Dividir a área em colunas
                    col1, col2 = st.columns(2)


                    with col1:
                        aprovador_liberado = st.text_input(f"Nome do aprovador para a solicitação {i+1} (Liberar)", key=f'aprovador_liberado_{i}')
                        if st.button(f"Liberar Chave", key=f"liberar_{i}"):
                            if aprovador_liberado:
                                update_status_aprovado(request['id'], "Liberada", aprovador_liberado)
                                st.success(f"Solicitação {request['id']} liberada.")
                                st.session_state['reset'] = True
                                time.sleep(2)
                                st.rerun()
                            else:
                                st.error("Insira o nome do aprovador.")

                    with col2:
                        aprovador_negado = st.text_input(f"Nome do aprovador para a solicitação {request['id']} (Negar)", key=f'aprovador_negado_{i}')
                        if st.button(f"Negar Chave", key=f"negar_{i}"):
                            if aprovador_negado:
                                update_status_negado(request['id'], "Negada", aprovador_negado)
                                st.warning(f"Solicitação {request['id']} negada.")
                                st.session_state['reset'] = True
                                time.sleep(2)
                                st.rerun()
                            else:
                                st.error("Insira o nome do aprovador.")
                        
            else:
                st.write("Nenhuma solicitação encontrada")

    with tab3:
        if not st.session_state.get("logged_in", False):
            login_screen_devolver_chaves()
        else:
            st.header("Devolução de Chaves")

            # Obter solicitações liberadas na aba devolução das chaves
            df_return = get_requests() # Essa função retorna todas as colicitações
            df_return = df_return[df_return['status'] == 'Liberada'] # Filtrar apenas as solicitações liberadas

            if not df_return.empty:
                for i, request in df_return.iterrows():
                    st.write(f"Solicitação: {request['id']}")
                    st.write(f"Data da Solicitação: {request['data_solicitacao']}")
                    st.write(f"Data de Retirada: {request['data_retirada']}")
                    st.write(f"Data de Entrega: {request['data_entrega']}")
                    st.write(f"Setor: {request['setor_solicitante']}")
                    st.write(f"Chave: {request['chave']}")
                    st.write(f"Solicitante: {request['email_solicitante']}")
                
                    recebedor = st.text_input(f"Nome do recebedor para a devolução da solicitação {i+1}", key=f'recebedor_{i}')
                
                    if st.button(f"Registrar Devolução para a solicitação {request['id']}", key=f"registrar_{i}"):
                        if recebedor:
                            register_return(request['id'], recebedor)
                            st.success(f"Devolução registrada para a solicitação {request['id']}.")
                            time.sleep(2)
                            st.rerun() # Atualiza a página
                        else:
                            st.error("Insira o nome do recebedor.")
            else:
                st.write("Nenhuma solicitação liberada encontrada.")

if __name__ == "__main__":
    main()