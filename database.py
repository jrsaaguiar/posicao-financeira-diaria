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
    data = Column(Date)
    empresa = Column(String)
    tipo_titulo = Column(String)
    valor = Column(Float)
    qtd_veiculos = Column(Integer, default=0)
    valor_medio = Column(Float, default=0.0)

# Cria a tabela se não existir
Base.metadata.create_all(bind=engine)
