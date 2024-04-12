import sqlalchemy as db
import settings as st
from logger import Logger

logger = Logger('db_helper', st.APPLICATION_LOG, write_to_stdout=st.DEBUG_MODE).get()


def trying_to_connect():
    conn = None
    try:
        params = st.DB_PARAMS

        engine = db.create_engine(
            f"postgresql+psycopg2://{params['user']}:{params['password']}@{params['host']}:5432/{params['database']}")

        logger.info('Connecting to the PostgreSQL database...')
        conn = engine.connect()
        logger.info('Connected to the PostgreSQL database')

        metadata = db.MetaData(schema="main")

        books = db.Table('books',
                         metadata,
                         db.Column('book_id', db.Integer, primary_key=True),
                         db.Column('book_name', db.Text),
                         db.Column('book_author', db.Text),
                         db.Column('book_year', db.Integer),
                         db.Column('book_is_taken', db.Boolean, default=False),
                         db.Column('abc', db.String(200))
                         )

        metadata.create_all(engine)

        insertion_query = books.insert().values([
            {'book_name': 'Бесы', 'book_author': 'Фёдор Достоевский', 'book_year': 1872, 'abc': '123'},
            {'book_name': 'Старик и море', 'book_author': 'Эрнест Хемингуэй', 'book_year': 1952, 'abc': None}
        ])

        conn.execute(insertion_query)
        conn.commit()

        select_all_query = db.select(books)
        select_all_results = conn.execute(select_all_query)

        for row in select_all_results.fetchall():
            print(row)
            print(row[0])

    except Exception as error:
        logger.error(f'Error: {error}')
    finally:
        if conn is not None:
            conn.close()
            logger.info('Database connection closed.')
