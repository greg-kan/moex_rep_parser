from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy import Column, Integer, String
from sqlalchemy import create_engine
from sqlalchemy.orm import Session
from sqlalchemy import select
from sqlalchemy import or_
from sqlalchemy import ForeignKey
from sqlalchemy.orm import relationship

import settings as st
from logger import Logger

SCHEMA_NAME = 'sql_alch_tutor'

Base = declarative_base()


class User(Base):
    __tablename__ = 'user'
    __table_args__ = {"schema": SCHEMA_NAME}

    def __init__(self, name, fullname):  # optional
        self.name = name
        self.fullname = fullname

    id = Column(Integer, primary_key=True)
    name = Column(String)
    fullname = Column(String)

    def __repr__(self):
        return "<User(%r, %r)>" % (self.name, self.fullname)


class Address(Base):
    __tablename__ = 'address'
    __table_args__ = {"schema": SCHEMA_NAME}

    id = Column(Integer, primary_key=True)
    email_address = Column(String, nullable=False)
    user_id = Column(Integer, ForeignKey(f'{SCHEMA_NAME}.user.id'))

    user = relationship("User", backref="addresses")

    def __repr__(self):
        return "<Address(%r)>" % self.email_address


def test_routine():
    print(User.__table__)
    print(User.__mapper__)

    fred_user = User(name='fred', fullname='Fred Jones')

    print(fred_user.name, fred_user.fullname)
    print(fred_user.id)

    params = st.DB_PARAMS
    engine = create_engine(
        f"postgresql+psycopg2://{params['user']}:{params['password']}@{params['host']}:5432/{params['database']}")

    Base.metadata.create_all(engine)

    session = Session(bind=engine)

    session.add(fred_user)

    # our_user = session.query(User).filter_by(name='ed').first()
    # print(our_user)
    # print(type(our_user))
    # print(ed_user.id)
    # print(our_user.id)
    # print(ed_user is our_user)
    #
    # session.add_all([User(name='wendy',
    #                       fullname='Wendy Weathersmith'),
    #                  User(name='mary',
    #                       fullname='Mary Contrary'),
    #                  User(name='fred',
    #                       fullname='Fred Flinstone')]
    #                 )
    #
    # print(session.dirty)
    # ed_user.fullname = 'Ed Jones'
    # print(session.dirty)
    # print(session.new)

    # session.commit()

    # ed_user.name = 'Edwardo'
    # fake_user = User(name='fakeuser', fullname='Invalid')
    # session.add(fake_user)
    #
    # results = session.query(User).filter(User.name.in_(['Edwardo', 'fakeuser'])).all()
    # print(results)
    #
    # session.rollback()
    #
    # print(ed_user.name)
    #
    # print(fake_user in session)
    #
    # results = session.query(User).filter(User.name.in_(['ed', 'fakeuser'])).all()
    # print(results)

###############################################

    # # sel = select([User.name, User.fullname]). \
    # #     where(User.name == 'ed'). \
    # #     order_by(User.id)
    # #
    # # result = session.connection().execute(sel).fetchall()
    # # print(result)
    #
    # query = session.query(User).filter(User.name == 'ed').order_by(User.id)
    # query.all()
    # for rec in query:
    #     print(rec)
    #
    # for name, fullname in session.query(User.name, User.fullname):
    #     print(name, fullname)
    #
    # for row in session.query(User, User.name):
    #     print(row.User, row.name)
    #
    # uu = session.query(User).order_by(User.id)[2]
    # print(uu)
    #
    # for u in session.query(User).order_by(User.id)[1:3]:
    #     print(u)
    #
    # for name in session.query(User.name).filter_by(fullname='Ed Jones'):
    #     print(name)
    #
    # for name in session.query(User.name).filter(User.fullname == 'Ed Jones'):
    #     print(name)
    #
    # for name, in session.query(User.name).filter(or_(User.fullname == 'Ed Jones', User.id < 5)):
    #     print(name)
    #
    # for user in session.query(User).filter(User.name == 'ed').filter(User.fullname == 'Ed Jones'):
    #     print(user)

##########################

    # query = session.query(User).filter_by(fullname='Ed Jones')
    # print(query.all())
    # print(query.first())
    # print(query.one())
    #
    # query = session.query(User).filter_by(fullname='nonexistent')
    # # print(query.one())
    #
    # try:
    #     query = session.query(User)
    #     query.one()
    # except Exception as e:
    #     print(str(e))

######################################

    jack = User(name='jack', fullname='Jack Bean')
    print(jack.addresses)

    jack.addresses = [Address(email_address='jack@gmail.com'),
                      Address(email_address='j25@yahoo.com'),
                      Address(email_address='jack@hotmail.com'), ]

    print(jack.addresses[1])
    print(jack.addresses[1].user)

    session.add(jack)

    print(session.new)

    session.commit()

    print(jack.addresses)
    print(jack.addresses)

    fred = session.query(User).filter_by(name='fred').one()
    jack.addresses[1].user = fred

    print(fred.addresses)

    session.commit()
