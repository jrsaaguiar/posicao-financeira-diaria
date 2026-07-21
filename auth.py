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
    
    # <- MUDOU AQUI: clear_on_submit=True e keys
    with st.form("form_cadastro", clear_on_submit=True):
        nome = st.text_input("Nome", key="cad_nome").upper()
        email = st.text_input("Email", key="cad_email").lower()
        senha = st.text_input("Senha", type="password", key="cad_senha")
        perfil = st.selectbox("Perfil", ["Usuario", "Admin"], key="cad_perfil")
        
        submitted = st.form_submit_button("Cadastrar")
        
        if submitted:
            if not nome or not email or not senha:
                st.error("Preencha todos os campos")
            else:
                ok, msg = criar_usuario(nome, email, senha, perfil)
                if ok: 
                    st.success(msg)
                    st.rerun() # <- MUDOU AQUI: recarrega e limpa a tela
                else: 
                    st.error(msg)