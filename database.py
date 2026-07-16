from sqlalchemy import create_engine, Column, Integer, String, Float, Date
from sqlalchemy.orm import sessionmaker, declarative_base
import streamlit as st

DATABASE_URL = st.secrets["DATABASE_URL"]

engine = create_engine(DATABASE_URL)
SessionLocal = sessionmaker(autocommit=False, autoflush=False, bind=engine)
Base = declarative_base()

class PosicaoDiaria(Base):
    __tablename__ = "posicoes_diarias"
    id = Column(Integer, primary_key=True, index=True)
    data = Column(String)  # <- MUDEI DE Date PARA String
    empresa = Column(String)
    tipo_titulo = Column(String)
    valor = Column(Float)
    qtd_veiculos = Column(Integer, default=0)
    valor_medio = Column(Float, default=0.0)

class Usuarios(Base):
    __tablename__ = 'usuarios'
    id = Column(Integer, primary_key=True)
    email = Column(String(100), unique=True, nullable=False)
    senha_hash = Column(String(255), nullable=False)
    nome = Column(String(100))
    ativo = Column(Boolean, default=True)
    
def init_db():
    Base.metadata.create_all(bind=engine)

init_db() # Cria a tabela assim que o arquivo é importado
