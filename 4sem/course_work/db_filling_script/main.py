import random
from datetime import timedelta, datetime
import psycopg2
from psycopg2.extras import DictCursor
import logging
from tqdm import tqdm
from faker import Faker

fake = Faker()

RENTS_COUNT = 300000
USERS_COUNT = 10000
CARS_COUNT = 500

# Database connection settings
DB_HOST = 'localhost'
DB_NAME = 'database_cw'
DB_USER = 'dmitry'
DB_PASSWORD = '1234'

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


def fff():
    logger.info("Connecting to the database...")
    conn = psycopg2.connect(
        host=DB_HOST,
        database=DB_NAME,
        user=DB_USER,
        password=DB_PASSWORD,
        cursor_factory=DictCursor
    )
    cur = conn.cursor()

    logger.info("Populating bank_data table...")
    for i in tqdm(range(1, USERS_COUNT + 1), desc="bank_data", unit="rows"):
        num = fake.credit_card_number(card_type='visa')
        date = fake.credit_card_expire()
        cvc = fake.random_int(100, 999)
        cur.execute("INSERT INTO bank_data"
                    "(id, num, date, cvc)"
                    " VALUES (%s,%s, %s, %s)",
                    (i, num, date, cvc))

    conn.commit()

    def generate_phone_number():
        code = random.choice(['79', '78', '76', '75', '74'])  # random operator code
        prefix = str(random.randint(100, 999))  # random prefix
        suffix = str(random.randint(10000, 99999))  # random suffix
        return f"{code}{prefix}{suffix}"

    def generate_email(_num):
        _num %= 3
        if _num == 0:
            return fake.free_email()
        if _num == 1:
            return fake.safe_email()
        if _num == 2:
            return fake.company_email()

    logger.info("Populating fine table...")
    for i in tqdm(range(1, 1001), desc="fine", unit="rows"):
        _sum = fake.random_int(100, 1000)
        description = fake.text()
        cur.execute("INSERT INTO fine "
                    "(id, sum, description) "
                    "VALUES (%s,%s, %s)",
                    (i, _sum, description))

    conn.commit()

    logger.info("Populating damage table...")
    for i in tqdm(range(1, 101), desc="damage", unit="rows"):
        machine_part = fake.random_element(['engine', 'transmission', 'brakes'])
        description = fake.text()
        severity = fake.random_int(1, 5)
        cur.execute("INSERT INTO damage "
                    "(id, machine_part, description, severity) "
                    "VALUES (%s,%s, %s, %s)",
                    (i, machine_part, description, severity))

    conn.commit()

    logger.info("Populating general_info table...")
    for i in tqdm(range(1, 101), desc="general_info", unit="rows"):
        checkup_date = fake.random_int(1, 12)
        engine_type = fake.random_element(['gasoline', 'diesel', 'electric'])
        color = fake.color_name()
        description = fake.text()
        cur.execute("INSERT INTO general_info"
                    "(id, checkup_date, engine_type, color, description)"
                    "VALUES (%s,%s, %s, %s, %s)",
                    (i, checkup_date, engine_type, color, description))

    conn.commit()

    logger.info("Populating transport_info table...")
    for i in tqdm(range(1, CARS_COUNT + 1), desc="transport_info", unit="rows"):
        brand = fake.company()[:32]  # truncate to 32 characters
        release_year = fake.random_int(2000, 2022)
        model = fake.sentence(nb_words=2)[:32]  # truncate to 32 characters
        license_level = fake.random_element(['A', 'B', 'C'])
        cur.execute("INSERT INTO transport_info"
                    "(id, brand, release_year, model, license_level) "
                    "VALUES (%s,%s, %s, %s, %s)",
                    (i, brand, release_year, model, license_level))

    conn.commit()

    logger.info("Populating transport table...")
    for i in tqdm(range(1, CARS_COUNT + 1), desc="transport", unit="rows"):
        general_info_id = fake.random_int(1, 100)
        cost_per_hour = fake.random_int(100, 1000)
        transport_info_id = i
        free = fake.boolean()
        state_number = fake.license_plate()
        today = datetime.today()
        year_ago = today - timedelta(days=365)

        date_add = fake.date_time_between(start_date=year_ago, end_date=today)
        cur.execute("INSERT INTO transport"
                    "(id, general_info_id, transport_info_id, free, state_number, date_add, cost_per_hour) "
                    "VALUES (%s,%s, %s, %s, %s, %s, %s)",
                    (i, general_info_id, transport_info_id, free, state_number, date_add, cost_per_hour))
    conn.commit()

    logger.info("Populating transport_damages table...")
    for i in tqdm(range(1, CARS_COUNT // 10 + 1), desc="transport_damages", unit="rows"):
        transport_id = fake.random_int(1, CARS_COUNT)
        damage_id = fake.random_int(1, 100)
        cur.execute("INSERT INTO transport_damages "
                    "(id, transport_id, damage_id) "
                    "VALUES (%s,%s, %s)",
                    (i, transport_id, damage_id))

    conn.commit()

    logger.info("Populating registration_address table...")
    for i in tqdm(range(1, USERS_COUNT + 1), desc="registration_address", unit="rows"):
        city = fake.city()
        street = fake.street_name()
        house_num = fake.building_number()
        flat = fake.random_int(1, 35)
        country = fake.country()
        cur.execute(
            "INSERT INTO registration_address"
            "(id, city, street, house_num, flat, country) "
            "VALUES (%s, %s, %s, %s, %s, %s)",
            (i, city, street, house_num, flat, country))

    conn.commit()

    logger.info("Populating human table...")
    for i in tqdm(range(1, USERS_COUNT + 1), desc="human", unit="rows"):
        surname = fake.last_name()
        name = fake.first_name()
        mid_name = fake.first_name() + "ovich"  # generate a middle name
        passport_series = fake.random_int(1000, 9999)
        passport_num = fake.random_int(100000, 999999)
        birthday = fake.date_of_birth(minimum_age=19, maximum_age=75)
        reg_address_id = i
        cur.execute("INSERT INTO human "
                    "(id, surname, name, mid_name, passport_series, passport_num, birthday, reg_address) "
                    "VALUES (%s, %s, %s, %s, %s, %s, %s, %s)",
                    (i, surname, name, mid_name, passport_series, passport_num, birthday, reg_address_id))

    conn.commit()

    logger.info("Populating user table...")
    for i in tqdm(range(1, USERS_COUNT + 1), desc="user", unit="rows"):
        nick = fake.user_name() + f"{i}"
        bank_data_id = i
        email = generate_email(i)
        phone_number = generate_phone_number()
        today = datetime.today()
        year_ago = today - timedelta(days=365)

        reg_date = fake.date_time_between(start_date=year_ago, end_date=today)
        human_id = i
        cur.execute("INSERT INTO \"user\" "
                    "(id, nick, bank_data, email, phone_number, reg_date, human_id) "
                    "VALUES (%s,%s, %s, %s, %s, %s, %s)",
                    (i, nick, bank_data_id, email, phone_number, reg_date, human_id))
    conn.commit()

    logger.info("Populating rent table...")
    for i in tqdm(range(1, RENTS_COUNT + 1), desc="rent", unit="rows"):
        transport_id = fake.random_int(1, CARS_COUNT)
        user_id = fake.random_int(1, USERS_COUNT)

        begin_date = fake.date_time()
        end_date = begin_date + timedelta(days=random.randint(1, 365))
        city = fake.city()
        cur.execute("INSERT INTO rent"
                    "(id, transport_id, user_id, begin_date, end_date, city) "
                    "VALUES (%s,%s, %s, %s, %s, %s)",
                    (i, transport_id, user_id, begin_date, end_date, city))

    conn.commit()

    logger.info("Populating cheque table...")
    for i in range(1, RENTS_COUNT + 1):
        rent_id = i
        user_id = fake.random_int(min=1, max=USERS_COUNT)
        _sum = fake.random_int(100, 10000)
        payment_status = fake.boolean()

        cur.execute("""
                INSERT INTO cheque (id, rent_id, user_id, total_cost, payment_status)
                VALUES (%s,%s, %s, %s, %s)
            """, (i, rent_id, user_id, _sum, payment_status))

    conn.commit()

    logger.info("Populating user_fines table...")
    for i in range(1, 1000):
        user_id = fake.random_int(1, USERS_COUNT)
        fine_id = fake.random_int(1, CARS_COUNT)

        today = datetime.today()
        year_ago = today - timedelta(days=365)

        date = fake.date_time_between(start_date=year_ago, end_date=today)
        cur.execute("INSERT INTO user_fines "
                    "(id, user_id, fine_id, date) "
                    "VALUES (%s,%s, %s, %s)",
                    (i, user_id, fine_id, date))

    conn.commit()

    logger.info("Populating driver_license table...")
    for i in range(1, USERS_COUNT + 1):
        _type = fake.random_element(['A', 'B', 'C'])
        begin_date = fake.date_time_between(start_date='-1y', end_date='now')
        end_date = begin_date + timedelta(days=random.randint(1, 365))
        num = fake.license_plate()
        cur.execute("INSERT INTO driver_license "
                    "(id, type, begin_date, end_date, num) "
                    "VALUES (%s, %s, %s, %s, %s)",
                    (i, _type, begin_date, end_date, num))

    conn.commit()

    logger.info("Populating user_license table...")
    for i in range(1, USERS_COUNT + 1):
        user_id = i
        licence_id = i
        cur.execute("INSERT INTO user_license "
                    "(id, user_id, licence_id) "
                    "VALUES (%s,%s, %s)",
                    (i, user_id, licence_id))

    conn.commit()

    logger.info("Done!")


fff()

