import streamlit as st
from sqlalchemy import create_engine, Column, String, Boolean, Integer, Date, Numeric, func # <-- ADICIONA TUDO AQUI
from sqlalchemy.orm import sessionmaker, declarative_base

DATABASE_URL = st.secrets["DATABASE_URL"]

engine = create_engine(DATABASE_URL, pool_pre_ping=True, pool_recycle=300)
SessionLocal = sessionmaker(autocommit=False, autoflush=False, bind=engine)
Base = declarative_base()

class Usuarios(Base):
    __tablename__ = "usuarios"
    
    email = Column(String(100), primary_key=True, index=True)
    senha_hash = Column(String(255))
    nome = Column(String(100))
    ativo = Column(Boolean, default=True)
    perfil = Column(String(20))

class PosicaoDiaria(Base):
    __tablename__ = "posicoes_diarias"
    
    id = Column(Integer, primary_key=True, index=True) # <-- agora reconhece Integer
    data = Column(Date, index=True) # <-- agora reconhece Date
    empresa = Column(String)
    tipo = Column(String, index=True)
    descricao = Column(String)
    valor = Column(Numeric(15, 2)) # <-- agora reconhece Numeric
    data_importacao = Column(Date, default=func.current_date()) # <-- agora reconhece func

def init_db():
    Base.metadata.create_all(bind=engine)
