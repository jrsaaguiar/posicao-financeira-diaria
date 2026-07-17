import streamlit as st
import hashlib
from database import SessionLocal, Usuarios

def hash_senha(senha):
    return hashlib.sha256(senha.encode()).hexdigest()

def verificar_login(email, senha):
    db = SessionLocal()
    usuario = db.query(Usuarios).filter(Usuarios.email == email).first()
    db.close()
    if usuario and usuario.senha_hash == hash_senha(senha) and usuario.ativo:
        return usuario
    return None

def criar_usuario(nome, email, senha, perfil="Usuario"):
    db = SessionLocal()
    if db.query(Usuarios).filter(Usuarios.email == email).first():
        db.close()
        return False, "Email já cadastrado"
    novo = Usuarios(
        nome=nome, 
        email=email, 
        senha_hash=hash_senha(senha),
        perfil=perfil
    )
    db.add(novo)
    db.commit()
    db.close()
    return True, "Usuário criado com sucesso"

def tela_cadastro_usuario():
    if st.session_state.get('perfil') != 'Admin':
        st.warning("Você não tem permissão para acessar esta tela.")
        return
    
    st.subheader("Cadastrar Novo Usuário")
    with st.form("form_cadastro"):
        nome = st.text_input("Nome")
        email = st.text_input("Email")
        senha = st.text_input("Senha", type="password")
        perfil = st.selectbox("Perfil", ["Usuario", "Admin"])
        if st.form_submit_button("Cadastrar"):
            ok, msg = criar_usuario(nome, email, senha, perfil)
            if ok: st.success(msg)
            else: st.error(msg)
