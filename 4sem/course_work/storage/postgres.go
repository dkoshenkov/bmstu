package storage

import (
	"context"
	"fmt"
	"github.com/goriiin/cw/config"
	"github.com/jackc/pgx/v5/pgxpool"
	"os"
)

type Storage struct {
	db *pgxpool.Pool
}

// New -
func New(pathToDB string) (*Storage, error) {
	const op = "db.postgres.New"

	cfg, err := config.ReadConfig(pathToDB)
	if err != nil {
		fmt.Println(fmt.Errorf("%s - config err: %w\n", op, err))
		os.Exit(1)
	}

	poolConfig, err := config.NewPoolConfig(cfg)
	if err != nil {
		fmt.Println(fmt.Errorf("%s - config err: %w\n", op, err))
		os.Exit(1)
	}

	poolConfig.MaxConns = 5

	conn, err := config.NewConnection(poolConfig)
	if err != nil {
		return nil, fmt.Errorf("%s: %w", op, err)
	}

	return &Storage{
		db: conn,
	}, nil
}

// TestSelect - вывод в консоль всех таблиц
func (s *Storage) TestSelect() error {
	const op = "db.postgres.TestSelect"

	rows, err := s.db.Query(context.Background(),
		`SELECT table_name
             FROM information_schema.tables 
             WHERE table_schema = 'public';`)

	if err != nil {
		return fmt.Errorf("%s - config err: %w", op, err)
	}

	for rows.Next() {
		var tableName string
		err = rows.Scan(&tableName)
		if err != nil {
			return fmt.Errorf("%s - config err: %w", op, err)
		}
		fmt.Printf("Created table name: %s\n", tableName)
	}

	// проверяем ошибку после обработки результатов
	if err = rows.Err(); err != nil {
		return fmt.Errorf("%s: %w", op, err)
	}
	return nil
}
