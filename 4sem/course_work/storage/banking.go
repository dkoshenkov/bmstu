package storage

import (
	"context"
	"fmt"
	"github.com/goriiin/cw/structures/bank"
	"github.com/goriiin/cw/structures/user"
)

func (s *Storage) FineSelect() ([]bank.Fine, error) {
	const op = "storage.banking.FineSelect"
	rows, err := s.db.Query(context.Background(),
		"select id, sum , description from fine")
	if err != nil {
		return nil, fmt.Errorf("%s - config err: %w", op, err)
	}
	defer rows.Close()
	fineSlice := make([]bank.Fine, 0)
	for rows.Next() {
		var fine bank.Fine
		err = rows.Scan(&fine.ID, &fine.Sum, &fine.Description)
		if err != nil {
			return nil, fmt.Errorf("%s - config err: %w", op, err)
		}
		fineSlice = append(fineSlice, fine)
	}
	return fineSlice, nil
}

func (s *Storage) ChequeSelect() ([]bank.Cheque, error) {
	const op = "storage.banking.ChequeSelect"
	rows, err := s.db.Query(context.Background(),
		"select id,rent_id, user_id,total_cost,payment_status  from cheque")
	if err != nil {
		return nil, fmt.Errorf("%s - config err: %w", op, err)
	}
	defer rows.Close()
	chequeSlice := make([]bank.Cheque, 0)
	for rows.Next() {
		var cheque bank.Cheque
		err = rows.Scan(&cheque.ID, &cheque.RentID, &cheque.UserID, &cheque.TotalCost, &cheque.PaymentStatus)
		if err != nil {
			return nil, fmt.Errorf("%s - config err: %w", op, err)
		}
		chequeSlice = append(chequeSlice, cheque)
	}
	return chequeSlice, nil
}

func (s *Storage) BankDataSelect() ([]bank.Data, error) {
	const op = "storage.banking.BankDataSelect"
	rows, err := s.db.Query(context.Background(),
		"select id, num, date, cvc from bank_data")
	if err != nil {
		return nil, fmt.Errorf("%s - config err: %w", op, err)
	}
	defer rows.Close()
	dataSlice := make([]bank.Data, 0)
	for rows.Next() {
		var data bank.Data
		err = rows.Scan(&data.ID, &data.Num, &data.Date, &data.CVC)
		if err != nil {
			return nil, fmt.Errorf("%s - config err: %w", op, err)
		}
		dataSlice = append(dataSlice, data)
	}
	return dataSlice, nil
}

func (s *Storage) UserFinesSelect() ([]bank.UserFines, error) {
	const op = "storage.banking.UserFinesSelect"
	rows, err := s.db.Query(context.Background(),
		"select id, user_id, fine_id, date from user_fines")
	if err != nil {
		return nil, fmt.Errorf("%s - config err: %w", op, err)
	}
	defer rows.Close()
	userFinesSlice := make([]bank.UserFines, 0)
	for rows.Next() {
		var userFines bank.UserFines
		err = rows.Scan(&userFines.ID, &userFines.UserID, &userFines.FineID, &userFines.Date)
		if err != nil {
			return nil, fmt.Errorf("%s - config err: %w", op, err)
		}
		userFinesSlice = append(userFinesSlice, userFines)
	}
	return userFinesSlice, nil
}

func (s *Storage) RentSelect() ([]user.Rent, error) {
	const op = "storage.banking.RentSelect"
	rows, err := s.db.Query(context.Background(),
		"select id, transport_id, user_id, begin_date, end_date, city from rent")
	if err != nil {
		return nil, fmt.Errorf("%s - config err: %w", op, err)
	}
	defer rows.Close()
	rentSlice := make([]user.Rent, 0)
	for rows.Next() {
		var rent user.Rent
		err = rows.Scan(&rent.ID, &rent.TransportID, &rent.UserID, &rent.BeginDate, &rent.EndDate, &rent.City)
		if err != nil {
			return nil, fmt.Errorf("%s - config err: %w", op, err)
		}
		rentSlice = append(rentSlice, rent)
	}
	return rentSlice, nil
}
