create table registration_address
(
    id        bigserial
        constraint id
            primary key,
    city      varchar(64) not null,
    street    varchar(64) not null,
    house_num varchar(10) not null,
    flat      integer     not null,
    country   varchar(64) not null
);

alter table registration_address
    owner to dmitry;

create table human
(
    surname         varchar(32)              not null,
    name            varchar(32)              not null,
    mid_name        varchar(32),
    passport_series smallint                 not null,
    passport_num    integer                  not null,
    birthday        timestamp with time zone not null,
    id              serial
        constraint human_id
            primary key,
    reg_address     integer                  not null
        constraint reg_address
            references registration_address
            on delete cascade,
    constraint passport
        unique (passport_num, passport_series)
);

alter table human
    owner to dmitry;

create table bank_data
(
    num  varchar(20) not null
        constraint bank_data_pk_2
            unique,
    date varchar(5)  not null
        constraint date_check
            check ((date)::text ~~ '__/__'::text),
    cvc  smallint    not null,
    id   bigserial
        constraint bank_data_pk
            primary key
);

alter table bank_data
    owner to dmitry;

create table "user"
(
    nick         varchar(50)              not null
        constraint nick
            unique,
    bank_data    integer                  not null
        constraint user_bank_data_id_fk
            references bank_data,
    email        varchar(100)             not null
        constraint check_name
            check ((email)::text ~~ '%@%'::text),
    phone_number varchar(20)              not null
        constraint phone_number
            unique,
    reg_date     timestamp with time zone not null,
    human_id     bigint                   not null
        constraint user_human_id_fk
            references human
            on update restrict on delete cascade,
    id           bigserial
        constraint user_pk
            primary key
);

alter table "user"
    owner to dmitry;

create table fine
(
    id          bigserial
        constraint fine_pk
            primary key,
    sum         money not null,
    description text
);

alter table fine
    owner to dmitry;

create table user_fines
(
    id      bigserial
        constraint user_fines_pk
            primary key,
    user_id bigint
        constraint user__fk
            references "user"
            on delete set null,
    fine_id bigint                   not null
        constraint fines___fk
            references fine,
    date    timestamp with time zone not null
);

alter table user_fines
    owner to dmitry;

create table driver_license
(
    id         bigserial
        constraint driver_license_pk
            primary key,
    type       varchar(5)               not null,
    begin_date timestamp with time zone not null,
    end_date   timestamp with time zone not null,
    num        varchar(10)              not null
);

alter table driver_license
    owner to dmitry;

create table user_license
(
    id         bigserial,
    user_id    bigint not null
        constraint user_license_user_id_fk
            references "user",
    licence_id bigint not null
        constraint user_license_driver_license_id_fk
            references driver_license
);

alter table user_license
    owner to dmitry;

create table transport_info
(
    id            bigserial
        constraint transport_info_pk
            primary key,
    brand         varchar(32) not null,
    release_year  smallint    not null,
    model         varchar(32) not null,
    license_level varchar(5)  not null
);

alter table transport_info
    owner to dmitry;

create table damage
(
    id           bigserial
        constraint damage_pk
            primary key,
    machine_part varchar(64) not null,
    description  text,
    severity     integer     not null
);

alter table damage
    owner to dmitry;

create table general_info
(
    id           bigserial
        constraint general_info_pk
            primary key,
    checkup_date integer     not null,
    engine_type  varchar(32) not null,
    color        varchar(32) not null,
    description  text
);

comment on column general_info.checkup_date is 'дата последнего тех-осмотра';

alter table general_info
    owner to dmitry;

create table transport
(
    id                bigserial
        constraint transport_pk
            primary key,
    general_info_id   bigint                   not null
        constraint transport_general_info_id_fk
            references general_info,
    transport_info_id bigint                   not null
        constraint transport___fk
            references transport_info,
    free              boolean                  not null,
    state_number      varchar(9)               not null,
    date_add          timestamp with time zone not null,
    cost_per_hour     integer default 300      not null
);

alter table transport
    owner to dmitry;

create table transport_damages
(
    id           bigserial
        constraint transport_damages_pk
            primary key,
    transport_id integer
        constraint transport_damages_transport_id_fk
            references transport,
    damage_id    integer
        constraint transport_damages_damage_id_fk
            references damage
);

alter table transport_damages
    owner to dmitry;

create table rent
(
    id           bigserial
        constraint rent_pk
            primary key,
    transport_id bigint      not null
        constraint rent_transport_id_fk
            references transport,
    user_id      bigint      not null
        constraint rent_user_id_fk
            references "user",
    begin_date   timestamp   not null,
    end_date     timestamp,
    city         varchar(32) not null
);

alter table rent
    owner to dmitry;

create table cheque
(
    id             bigserial
        constraint cheque_pk
            primary key,
    rent_id        bigint  not null
        constraint cheque___fk
            references rent,
    user_id        bigint  not null
        constraint cheque_user_id_fk
            references "user",
    total_cost     money   not null,
    payment_status boolean not null
);

alter table cheque
    owner to dmitry;

-- Additional logical constraints and triggers

-- Check age constraint
alter table human
    add constraint age_check
        check (extract(year from age(birthday)) > 18);

-- Create trigger to check age
create or replace function check_age()
    returns trigger as
$$
begin
    if extract(year from age(NEW.birthday)) <= 18 then
        raise 'Age must be greater than 18';
    end if;
    return NEW;
end;
$$
    language plpgsql;

create trigger check_age_trigger
    before insert or update
    on human
    for each row
execute procedure check_age();

-- Check phone number format constraint
alter table "user"
    add constraint phone_number_format_check
        check (phone_number ~ '^\d{10}$');

-- Create trigger to check phone number format
create or replace function check_phone_number()
    returns trigger as
$$
begin
    if NEW.phone_number !~ '^\d{10}$' then
        raise 'Invalid phone number format';
    end if;
    return NEW;
end;
$$
    language plpgsql;

create trigger check_phone_number_trigger
    before insert or update
    on "user"
    for each row
execute procedure check_phone_number();

-- Check email format constraint
alter table "user"
    add constraint email_format_check
        check (email ~ '^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$');

-- Create trigger to check email format
create or replace function check_email()
    returns trigger as
$$
begin
    if NEW.email !~ '^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$' then
        raise 'Invalid email format';
    end if;
    return NEW;
end;
$$
    language plpgsql;

create trigger check_email_trigger
    before insert or update
    on "user"
    for each row
execute procedure check_email();

-- Check bank data format constraint
alter table bank_data
    add constraint card_number_format_check
        check (num ~ '^\d{16}$');

-- Create trigger to check bank data format
create or replace function check_bank_data()
    returns trigger as
$$
begin
    if NEW.num !~ '^\d{16}$' then
        raise 'Invalid bank data format';
    end if;
    return NEW;
end;
$$
    language plpgsql;

create trigger check_bank_data_trigger
    before insert or update
    on bank_data
    for each row
execute procedure check_bank_data();

