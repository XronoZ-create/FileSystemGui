from sqlalchemy import Column, Integer, String
from database import Base

class Svazi(Base):
    __tablename__ = 'svazi'
    id = Column(Integer, primary_key=True)
    main = Column(String)
    second = Column(String)
