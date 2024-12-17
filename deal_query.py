import pandas as pd
import os
from sqlalchemy import create_engine


# Define the connection string
connection_string = "postgresql+psycopg2://kyleanderson:your_password@pacs-azure-psql.postgres.database.azure.com:5432/acquisition"


class DatabaseFetcher:
    def __init__(self, connection_string):
        """
        Initialize the DatabaseFetcher with a connection string.

        Args:
            connection_string (str): Database connection string.
        """
        self.connection_string = connection_string
        self.engine = None

    def connect(self):
        """
        Create a database connection using SQLAlchemy.
        """
        if not self.connection_string:
            raise ValueError("Connection string is empty. Please provide a valid connection string.")
        self.engine = create_engine(self.connection_string)

    def fetch_data(self, query, chunksize=10000):
        """
        Fetch data from the database in chunks.

        Args:
            query (str): SQL query to execute.
            chunksize (int): Number of rows to fetch per chunk.

        Returns:
            pd.DataFrame: Concatenated DataFrame containing all the fetched data.
        """
        if not self.engine:
            raise ConnectionError("Database connection is not established. Call connect() first.")

        chunks = []
        chunk_count = 1
        for chunk in pd.read_sql(query, self.engine, chunksize=chunksize):
            chunks.append(chunk)
            print(f"Fetched chunk {chunk_count}")
            chunk_count += 1

        # Combine all chunks into a single DataFrame
        return pd.concat(chunks, ignore_index=True)

    def close_connection(self):
        """
        Close the database connection.
        """
        if self.engine:
            self.engine.dispose()
            self.engine = None
            print("Database connection closed.")
