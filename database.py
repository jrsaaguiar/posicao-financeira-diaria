# database.py
import streamlit as st
from sqlalchemy import create_engine, Column, String, Boolean, Integer, Date, Numeric, Float, func
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
    
    id = Column(Integer, primary_key=True, index=True)
    data = Column(Date, index=True)
    empresa = Column(String)
    tipo_titulo = Column(String, index=True) # <-- CORRIGIDO: mudado de 'tipo' para 'tipo_titulo'
    valor = Column(Numeric(15, 2))
    qtd_veiculos = Column(Integer, default=0) # <-- Adicionado
    valor_medio = Column(Float, default=0.0)   # <-- Adicionado
    criado_por = Column(String, nullable=True) # <-- Adicionado
    data_importacao = Column(Date, default=func.current_date())

def init_db():
    Base.metadata.create_all(bind=engine)
