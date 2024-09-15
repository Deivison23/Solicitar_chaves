import streamlit as st
import pandas as pd
from datetime import datetime
import os
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

# Carregar ou criar o arquivo Excel
def load_excel(file_path):
    try:
        df = pd.read_excel(file_path)
        # Adiciona a coluna "Atraso" se não existir
        if "Atraso" not in df.columns:
            df["Atraso"] = ""
    except FileNotFoundError:
        df = pd.DataFrame(columns=["Data Solicitação","Data Retirada", "Data Entrega", "Setor Solicitante", "Chave", "Status", "Email Solicitante", "Nome aprovador chave", "Nome recebedor chave", "Atraso", "Data ultimo email"])
        df.to_excel(file_path, index=False)
    return df

# Função para adicionar nova solicitação
def add_request(file_path, data_retirada, data_entrega, setor_solicitante, email, chave):
    df = load_excel(file_path)

    # Criar um novo DataFrame com os dados da nova solicitação
    nova_solicitacao = pd.DataFrame({
        "Data Solicitação": [datetime.now().strftime("%d-%m-%Y %H:%M:%S")],
        "Data Retirada": [data_retirada],
        "Data Entrega": [data_entrega],
        "Setor Solicitante": [setor_solicitante],
        "Email Solicitante": [email],
        "Chave": [chave],
        "Status": ["Pendente"],
        "Nome aprovador chave": [""], #Inicialmente vazio
        "Nome recebedor chave": [""]
    })
    # Concatenar o novo DataFrame ao DataFrame existente
    df = pd.concat([df, nova_solicitacao], ignore_index=True)
    df.to_excel(file_path, index=False)

    # Enviar e-mail para o solicitante
    subject = "Confirmação de Solicitação de Chave"
    body = f"""
    Olá,

    Sua solicitação de chave foi recebida.

    Data da Solicitação: {nova_solicitacao["Data Solicitação"].values[0]}
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

    Data da Solicitação: {nova_solicitacao["Data Solicitação"].values[0]}
    Data de Retirada: {data_retirada}
    Data de Entrega: {data_entrega}
    Setor Solicitante: {setor_solicitante}
    Chave: {chave}
    Solicitante: {email}
    """
    send_email(ADMIN_EMAILS, subject_admin, body_admin)

def check_and_notify_delays(file_path):
    df = load_excel(file_path)    
    hoje = datetime.now().date()
    
    for index, row in df.iterrows():
        if row["Status"] == "Liberada":
            data_entrega = pd.to_datetime(row["Data Entrega"], format="%d-%m-%Y").date()
            
            if data_entrega < hoje:
                # Calcular dias de atraso
                atraso_atual = (hoje - data_entrega).days
                
                # Atualizar a coluna "Atraso"
                if pd.isna(row["Atraso"]) or int(row["Atraso"]) < atraso_atual:
                    df.at[index, "Atraso"] = atraso_atual

                    # Verificar a data do último email
                    ultima_data_email = row.get("Data ultimo email")
                    if pd.isna(ultima_data_email) or pd.to_datetime(ultima_data_email).date() < hoje:          
                    
                        # Enviar notificação de atraso
                        email = row["Email Solicitante"]
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
                        df.at[index, "Data Último E-mail"] = hoje.strftime("%d-%m-%Y")
                        df.to_excel(file_path, index=False)
    
# Função para atualizar o status da solicitação APROVADA
def update_status_aprovado(file_path, index, status, aprovador_liberado):
    df = load_excel(file_path)
    df.at[index, "Status"] = status
    df.at[index, "Nome aprovador chave"] = aprovador_liberado
    df.to_excel(file_path, index=False)

    email = df.at[index, "Email Solicitante"]
    subject = "Atualização sobre sua Solicitação de Chave"
    body = f"""
    Olá,

    Sua solicitação de chave foi aprovada.

    Data da Solicitação: {df.at[index, "Data Solicitação"]}
    Data de Retirada: {df.at[index, "Data Retirada"]}
    Data de Entrega: {df.at[index, "Data Entrega"]}
    Setor Solicitante: {df.at[index, "Setor Solicitante"]}
    Chave: {df.at[index, "Chave"]}
    Aprovador: {aprovador_liberado}

    Se apresentar na sala Operacional para retirada da chave com o aprovador.

    Atenciosamente,
    SOSC - Sistema Operacional de Solicitação de Chaves
    """
    send_email(email, subject, body)

# Função para atualizar o status da solicitação NEGADA
def update_status_negado(file_path, index, status, aprovador_negado):
    df = load_excel(file_path)
    df.at[index, "Status"] = status
    df.at[index, "Nome aprovador chave"] = aprovador_negado
    df.to_excel(file_path, index=False)

    email = df.at[index, "Email Solicitante"]
    subject = "Atualização sobre sua Solicitação de Chave"
    body = f"""
    Olá,

    Sua solicitação de chave foi Negada.

    Data da Solicitação: {df.at[index, "Data Solicitação"]}
    Data de Retirada: {df.at[index, "Data Retirada"]}
    Data de Entrega: {df.at[index, "Data Entrega"]}
    Setor Solicitante: {df.at[index, "Setor Solicitante"]}
    Chave: {df.at[index, "Chave"]}
    Aprovador: {aprovador_negado}

    No momento a chave não se encontra na sala operacional.

    Atenciosamente,
    SOSC - Sistema Operacional de Solicitação de Chaves
    """
    send_email(email, subject, body)

# Função para registrar a devolução da chave
def register_return(file_path, index, recebedor):
    df = load_excel(file_path)
    df.at[index, "Status"] = "Devolvida"
    df.at[index, "Nome recebedor chave"] = recebedor
    df.to_excel(file_path, index=False)

    st.success(f"Chave devolvida registrada com sucesso. Nome do recebedor: {recebedor}")


# Tela de Login liberação das chaves
def login_screen_liberar_chaves():
    st.title("Tela de Login")
    senha = st.text_input("Digite a senha:", type="password", key="senha_login_liberar_chaves")
    
    if st.button("Entrar", key="entrar_liberar"):
        if senha == SENHA_ACESSO:
            st.session_state["logged_in"] = True
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
            st.rerun()
        else:
            st.error("Senha incorreta. Tente novamente.")



# Interface do aplicativo
def main():
    # Verificar e notificar sobre atrasos
    check_and_notify_delays("solicitacoes_chaves.xlsx")
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
            if email and setor and data_retirada and data_entrega and chave:
                # Converter as datas para strings no formato correto
                data_retirada_str = data_retirada.strftime("%d-%m-%Y")
                data_entrega_str = data_entrega.strftime("%d-%m-%Y")    

                add_request(excel_file, data_retirada_str, data_entrega_str, setor, email, chave)
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
            pending_requests = pd.DataFrame()

            try:
                df = load_excel(excel_file)
                if not df.empty:
                    pending_requests = df[df["Status"] == "Pendente"]

            except Exception as e:
                st.error(f"Erro ao carregar o arquivo Excel: {e}")

            if not pending_requests.empty:
                for i, request in pending_requests.iterrows():
                    st.write(f"Solicitação {i+1}:")
                    st.write(f"Data de Retirada: {request['Data Retirada']}")
                    st.write(f"Data de Entrega: {request['Data Entrega']}")
                    st.write(f"Setor: {request['Setor Solicitante']}")
                    st.write(f"Chave: {request['Chave']}")
                    st.write(f"Solicitante: {request['Email Solicitante']}")
                        

                    # Dividir a área em colunas
                    col1, col2 = st.columns(2)


                    with col1:
                        aprovador_liberado = st.text_input(f"Nome do aprovador para a solicitação {i+1} (Liberar)", key=f'aprovador_liberado_{i}')
                        if st.button(f"Liberar Chave", key=f"liberar_{i}"):
                            if aprovador_liberado:
                                update_status_aprovado(excel_file, i, "Liberada", aprovador_liberado)
                                st.success(f"Solicitação {i+1} liberada.")
                                st.session_state['reset'] = True
                                time.sleep(2)
                                st.rerun()
                            else:
                                st.error("Insira o nome do aprovador.")

                    with col2:
                        aprovador_negado = st.text_input(f"Nome do aprovador para a solicitação {i+1} (Negar)", key=f'aprovador_negado_{i}')
                        if st.button(f"Negar Chave", key=f"negar_{i}"):
                            if aprovador_negado:
                                update_status_negado(excel_file, i, "Negada", aprovador_negado)
                                st.warning(f"Solicitação {i+1} negada.")
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

            # Inicializar df_return como um DataFrame vazio
            df_return = pd.DataFrame(columns=["Data Solicitação", "Data Retirada", "Data Entrega", "Setor Solicitante", "Chave", "Status", "Email Solicitante", "Nome aprovador chave", "Nome recebedor chave"])
            
            try:
                # tente carregar o arquivo Excel e filtrar as solicitações liberadas
                df_return = pd.read_excel(excel_file)
                df_return = df_return[df_return["Status"] == "Liberada"]
            except FileNotFoundError:
                st.warning("Arquivo de solicitações não encontrado.")
            except Exception as e:
                st.error("Erro ao carregar o arquivo Excel: {e}")

            if not df_return.empty:
                for i, request in df_return.iterrows():
                    st.write(f"Solicitação {i+1}:")
                    st.write(f"Data da Solicitação: {request['Data Solicitação']}")
                    st.write(f"Data de Retirada: {request['Data Retirada']}")
                    st.write(f"Data de Entrega: {request['Data Entrega']}")
                    st.write(f"Setor: {request['Setor Solicitante']}")
                    st.write(f"Chave: {request['Chave']}")
                    st.write(f"Solicitante: {request['Email Solicitante']}")
                
                    recebedor = st.text_input(f"Nome do recebedor para a devolução da solicitação {i+1}", key=f'recebedor_{i}')
                
                    if st.button(f"Registrar Devolução para a solicitação {i+1}", key=f"registrar_{i}"):
                        if recebedor:
                            register_return(excel_file, i, recebedor)
                            st.success(f"Devolução registrada para a solicitação {i+1}.")
                            time.sleep(2)
                            st.rerun()
                        else:
                            st.error("Insira o nome do recebedor.")
            else:
                st.write("Nenhuma solicitação liberada encontrada.")

if __name__ == "__main__":
    main()