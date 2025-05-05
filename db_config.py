# db_config.py
from sqlalchemy import create_engine

# substitua com a sua string real
DATABASE_URL = "postgresql://user:password@host:port/dbname"

engine = create_engine(DATABASE_URL)
