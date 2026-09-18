package config

import (
	"context"
	"fmt"
	"github.com/ilyakaznacheev/cleanenv"
	"github.com/jackc/pgx/v5/pgxpool"
	"gopkg.in/yaml.v3"
	"log"
	"net/url"
	"os"
	"time"
)

type DbConfig struct {
	Host     string `yaml:"host"`
	Port     string `yaml:"port"`
	Username string `yaml:"username"`
	Password string `yaml:"password"`
	DbName   string `yaml:"db_name"`
	Timeout  int    `yaml:"timeout"`
}

type LocalConfig struct {
	Env          string `yaml:"env"`
	DbConfigPath string `yaml:"db_conf_path"`
	HTTPServer   `yaml:"http_server"`
}

type HTTPServer struct {
	Address     string        `yaml:"address"`
	Timeout     time.Duration `yaml:"timeout"`
	IdleTimeout time.Duration `yaml:"idle_timeout"`
}

func MustLoad() *LocalConfig {
	configPath := "./config/server_config.yaml"

	if _, err := os.Stat(configPath); os.IsNotExist(err) {
		log.Fatal("config file not found")
	}

	var cfg LocalConfig
	if err := cleanenv.ReadConfig(configPath, &cfg); err != nil {
		log.Fatalf("cannot read config: %s", err)
	}
	return &cfg
}

func ReadConfig(pathToDB string) (*DbConfig, error) {
	yFile, err := os.ReadFile(pathToDB)
	if err != nil {
		log.Fatalf("cannot read config: %s", err)
		return nil, err
	}

	var cfg DbConfig
	_ = yaml.Unmarshal(yFile, &cfg)
	if err != nil {
		log.Fatalf("cannot unmarshaling: %s", err)
		return nil, err
	}

	return &cfg, nil
}

func NewPoolConfig(cfg *DbConfig) (*pgxpool.Config, error) {
	connStr := fmt.Sprintf("%s://%s:%s@%s:%s/%s?sslmode=disable&connect_timeout=%d",
		"postgres",
		url.QueryEscape(cfg.Username),
		url.QueryEscape(cfg.Password),
		cfg.Host,
		cfg.Port,
		cfg.DbName,
		cfg.Timeout)

	poolConfig, err := pgxpool.ParseConfig(connStr)
	if err != nil {
		log.Fatalf("cannot parse config: %s", err)
		return nil, err
	}

	return poolConfig, nil
}

func NewConnection(poolConfig *pgxpool.Config) (*pgxpool.Pool, error) {
	conn, err := pgxpool.NewWithConfig(context.Background(), poolConfig)
	if err != nil {
		return nil, err
	}

	return conn, nil
}

func SetupENV() {
	err := os.Setenv("CONFIG_PATH", "./config/local.yaml")
	if err != nil {
		fmt.Println("err: ", err)
	}
}
