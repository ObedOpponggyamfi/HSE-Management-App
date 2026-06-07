"""Local-first HSE dashboard package."""

from .database import DEFAULT_DB_PATH, make_engine, initialize_database

__all__ = ["DEFAULT_DB_PATH", "make_engine", "initialize_database"]
