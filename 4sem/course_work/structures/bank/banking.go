package bank

import (
	"time"
)

type Cheque struct {
	ID            int64  `db:"id"`
	RentID        int64  `db:"rent_id"`
	UserID        int64  `db:"user_id"`
	TotalCost     string `db:"total_cost"`
	PaymentStatus bool   `db:"payment_status"`
}

type Fine struct {
	ID          int64  `db:"id" json:"id"`
	Sum         string `db:"sum" json:"sum"`
	Description string `db:"description" json:"description"`
}

type UserFines struct {
	ID     int64     `db:"id"`
	UserID int64     `db:"user_id"`
	FineID int64     `db:"fine_id"`
	Date   time.Time `db:"date"`
}

type Data struct {
	ID   int64  `db:"id"`
	Num  string `db:"num"`
	Date string `db:"date"`
	CVC  int    `db:"cvc"`
}
