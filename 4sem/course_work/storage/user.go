package storage

import (
	"context"
	"fmt"
	"github.com/goriiin/cw/structures/user"
)

func (s *Storage) HumanSelect() ([]user.Human, error) {
	const op = "db.postgres.HumanSelect"
	rows, err := s.db.Query(context.Background(),
		"SELECT id, surname, name, mid_name, passport_series, passport_num, birthday, reg_address FROM human")
	if err != nil {
		return nil, fmt.Errorf("%s - config err: %w", op, err)
	}
	defer rows.Close()

	humans := make([]user.Human, 0)

	for rows.Next() {
		var human user.Human
		err := rows.Scan(&human.ID, &human.Surname, &human.Name, &human.MidName, &human.PassportSeries, &human.PassportNum, &human.Birthday, &human.RegAddress)
		if err != nil {
			return nil, fmt.Errorf("%s - config err: %w", op, err)
		}
		humans = append(humans, human)
	}

	return humans, nil
}

func (s *Storage) UserSelect() ([]user.User, error) {
	const op = "db.postgres.UserSelect"
	rows, err := s.db.Query(context.Background(),
		"select nick, bank_data, email, phone_number, reg_date, human_id, id from \"user\"")
	if err != nil {
		return nil, fmt.Errorf("%s - config err: %w", op, err)
	}
	defer rows.Close()
	users := make([]user.User, 0)
	for rows.Next() {
		var u user.User
		err = rows.Scan(&u.Nick, &u.BankData, &u.Email, &u.PhoneNumber, &u.RegDate, &u.HumanID, &u.ID)
		if err != nil {
			return nil, fmt.Errorf("%s - config err: %w", op, err)
		}
		users = append(users, u)
	}
	return users, nil
}

func (s *Storage) UserLicenceSelect() ([]user.License, error) {
	const op = "storage.user.UserLicenceSelect"
	rows, err := s.db.Query(context.Background(),
		"select id, user_id, licence_id from user_license")
	if err != nil {
		return nil, fmt.Errorf("%s - config err: %w", op, err)
	}
	defer rows.Close()
	licenses := make([]user.License, 0)
	for rows.Next() {
		var l user.License
		err = rows.Scan(&l.ID, &l.UserID, &l.LicenseID)
		if err != nil {
			return nil, fmt.Errorf("%s - config err: %w", op, err)
		}
		licenses = append(licenses, l)
	}
	return licenses, nil
}

func (s *Storage) RegistrationAddressSelect() ([]user.RegistrationAddress, error) {
	const op = "storage.user.RegistrationAddressSelect"
	rows, err := s.db.Query(context.Background(),
		"select id, city, street, house_num, flat, country from registration_address")
	if err != nil {
		return nil, fmt.Errorf("%s - config err: %w", op, err)
	}
	defer rows.Close()
	registrationAddresses := make([]user.RegistrationAddress, 0)
	for rows.Next() {
		var r user.RegistrationAddress
		err = rows.Scan(&r.ID, &r.City, &r.Street, &r.HouseNum, &r.Flat, &r.Country)
		if err != nil {
			return nil, fmt.Errorf("%s - config err: %w", op, err)
		}
		registrationAddresses = append(registrationAddresses, r)
	}
	return registrationAddresses, err
}
