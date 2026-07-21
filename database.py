import streamlit as st
from sqlalchemy import create_engine, Column, Integer, String, Date, Numeric, func # <-- ADICIONA Column e Integer aqui
from sqlalchemy.orm import sessionmaker, declarative_base

DATABASE_URL = st.secrets["DATABASE_URL"]

engine = create_engine(
    DATABASE_URL,
    pool_pre_ping=True,
    pool_recycle=300
)

SessionLocal = sessionmaker(autocommit=False, autoflush=False, bind=engine)
Base = declarative_base()

class PosicaoDiaria(Base):
    __tablename__ = "posicoes_diarias"
    
    id = Column(Integer, primary_key=True, index=True) # <-- agora ele vai reconhecer
    data = Column(Date, index=True)
    empresa = Column(String)
    tipo = Column(String, index=True)
    descricao = Column(String)
    valor = Column(Numeric(15, 2))
    data_importacao = Column(Date, default=func.current_date())

class Usuarios(Base):
    __tablename__ = "usuarios"
    
    id = Column(Integer, primary_key=True, index=True)
    nome = Column(String)
    email = Column(String, unique=True, index=True)
    senha_hash = Column(String)
    tipo = Column(String, default="usuario")

def init_db():
    Base.metadata.create_all(bind=engine)
