package storage

import (
	"context"
	"fmt"
	"github.com/goriiin/cw/structures/transport"
	"github.com/goriiin/cw/structures/user"
)

func (s *Storage) GeneralInfoSelect() ([]transport.GeneralInfo, error) {
	const op = "storage.transport.GeneralInfoSelect"
	rows, err := s.db.Query(context.Background(),
		`select id, checkup_date, engine_type, color,description from general_info`)
	if err != nil {
		return nil, fmt.Errorf("%s - config err: %w", op, err)
	}
	defer rows.Close()
	ginfoSlice := make([]transport.GeneralInfo, 0)
	for rows.Next() {
		var info transport.GeneralInfo
		err = rows.Scan(&info.ID, &info.CheckupDate, &info.EngineType, &info.Color, &info.Description)
		if err != nil {
			return nil, fmt.Errorf("%s - config err: %w", op, err)
		}
		ginfoSlice = append(ginfoSlice, info)
	}
	return ginfoSlice, nil
}

func (s *Storage) DriverLicenceSelect() ([]user.DriverLicense, error) {
	const op = "storage.transport.DriverLicenceSelect"
	rows, err := s.db.Query(context.Background(),
		"select id, type,begin_date, end_date, num from driver_license")
	if err != nil {
		return nil, fmt.Errorf("%s - config err: %w", op, err)
	}
	defer rows.Close()
	dLicenses := make([]user.DriverLicense, 0)
	for rows.Next() {
		var dLicense user.DriverLicense
		err = rows.Scan(&dLicense.ID, &dLicense.Type, &dLicense.BeginDate, &dLicense.EndDate, &dLicense.Num)
		if err != nil {
			return nil, fmt.Errorf("%s - config err: %w", op, err)
		}
		dLicenses = append(dLicenses, dLicense)
	}
	return dLicenses, nil
}

func (s *Storage) DamageSelect() ([]transport.Damage, error) {
	const op = "storage.transport.DamageSelect"
	rows, err := s.db.Query(context.Background(),
		"select id,machine_part,severity,description from damage")
	if err != nil {
		return nil, fmt.Errorf("%s - config err: %w", op, err)
	}
	defer rows.Close()
	damages := make([]transport.Damage, 0)
	for rows.Next() {
		var damage transport.Damage
		err = rows.Scan(&damage.ID, &damage.MachinePart, &damage.Severity, &damage.Description)
		if err != nil {
			return nil, fmt.Errorf("%s - config err: %w", op, err)
		}
		damages = append(damages, damage)
	}
	return damages, nil
}

func (s *Storage) TransportSelect() ([]transport.Transport, error) {
	const op = "storage.transport.TransportSelect"
	rows, err := s.db.Query(context.Background(),
		"select id, general_info_id, transport_info_id, free, state_number, date_add from transport")
	if err != nil {
		return nil, fmt.Errorf("%s - config err: %w", op, err)
	}
	defer rows.Close()
	transports := make([]transport.Transport, 0)
	for rows.Next() {
		var t transport.Transport
		err = rows.Scan(&t.ID, &t.GeneralInfoID, &t.TransportInfoID, &t.Free, &t.StateNumber, &t.DateAdd)
		if err != nil {
			return nil, fmt.Errorf("%s - config err: %w", op, err)
		}
		transports = append(transports, t)
	}
	return transports, nil
}

func (s *Storage) TransportDamagesSelect() ([]transport.Damages, error) {
	const op = "storage.transport.TransportDamagesSelect"
	rows, err := s.db.Query(context.Background(),
		"select id, transport_id, damage_id from transport_damages")
	if err != nil {
		return nil, fmt.Errorf("%s - config err: %w", op, err)
	}
	defer rows.Close()
	damages := make([]transport.Damages, 0)
	for rows.Next() {
		var d transport.Damages
		err = rows.Scan(&d.ID, &d.TransportID, &d.DamageID)
		if err != nil {
			return nil, fmt.Errorf("%s - config err: %w", op, err)
		}
		damages = append(damages, d)
	}
	return damages, nil
}

func (s *Storage) TransportInfoSelect() ([]transport.Info, error) {
	const op = "storage.transport.TransportInfoSelect"
	rows, err := s.db.Query(context.Background(),
		"select id, brand, release_year, model, license_level from transport_info")
	if err != nil {
		return nil, fmt.Errorf("%s - config err: %w", op, err)
	}
	defer rows.Close()
	infos := make([]transport.Info, 0)
	for rows.Next() {
		var i transport.Info
		err = rows.Scan(&i.ID, &i.Brand, &i.ReleaseYear, &i.Model, &i.LicenseLevel)
		if err != nil {
			return nil, fmt.Errorf("%s - config err: %w", op, err)
		}
		infos = append(infos, i)
	}
	return infos, nil
}
