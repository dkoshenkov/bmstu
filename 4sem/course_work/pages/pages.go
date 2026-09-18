package pages

import (
	"github.com/goriiin/cw/storage"
	"github.com/goriiin/cw/structures/bank"
	"github.com/goriiin/cw/structures/transport"
	"github.com/goriiin/cw/structures/user"
)

type BasePage struct {
	Content string
	Fields  []string
	Data    interface{}
}

func InitBasePage() BasePage {
	return BasePage{
		Content: "КУРСОВАЯ РАБОТА -- ВЫПОЛНИЛ КОШЕНКОВ Д.О",
	}
}

type BankData struct {
	Content string
	Fields  []string
	Data    []bank.Data
}

func InitBankData(s *storage.Storage) (BankData, error) {
	data, err := s.BankDataSelect()
	if err != nil {
		return BankData{}, err
	}

	return BankData{
		Content: "Банковская информация",
		Fields:  []string{"id", "num", "date", "cvc"},
		Data:    data,
	}, nil
}

type Cheque struct {
	Content string
	Fields  []string
	Data    []bank.Cheque
}

func InitCheque(s *storage.Storage) (Cheque, error) {
	data, err := s.ChequeSelect()
	if err != nil {
		return Cheque{}, err
	}
	return Cheque{
		Content: "Чеки",
		Fields:  []string{"id", "rent_id", "user_id", "total_cost", "payment_status"},
		Data:    data,
	}, nil
}

type Damage struct {
	Content string
	Fields  []string
	Data    []transport.Damage
}

func InitDamage(s *storage.Storage) (Damage, error) {
	data, err := s.DamageSelect()
	if err != nil {
		return Damage{}, err
	}
	return Damage{
		Content: "Виды Урона",
		Data:    data,
		Fields:  []string{"id", "machine_part", "description", "severity"},
	}, nil
}

type DriverLicense struct {
	Content string
	Fields  []string
	Data    []user.DriverLicense
}

func InitDriverLicense(s *storage.Storage) (DriverLicense, error) {
	data, err := s.DriverLicenceSelect()
	if err != nil {
		return DriverLicense{}, err
	}
	return DriverLicense{
		Content: "Водительские права",
		Data:    data,
		Fields:  []string{"id", "type", "begin_date", "end_date", "num"},
	}, nil
}

type Fine struct {
	Content string
	Fields  []string
	Data    []bank.Fine
}

func InitFine(s *storage.Storage) (Fine, error) {
	data, err := s.FineSelect()
	if err != nil {
		return Fine{}, err
	}
	return Fine{
		Content: "Штрафы",
		Data:    data,
		Fields:  []string{"id", "sum", "description"},
	}, nil
}

type GeneralInfo struct {
	Content string
	Fields  []string
	Data    []transport.GeneralInfo
}

func InitGeneralInfo(s *storage.Storage) (GeneralInfo, error) {
	data, err := s.GeneralInfoSelect()
	if err != nil {
		return GeneralInfo{}, err
	}
	return GeneralInfo{
		Content: "Общая информация о машине",
		Fields:  []string{"id", "checkup_date", "engine_type", "color", "description"},
		Data:    data,
	}, nil
}

type Human struct {
	Content string
	Fields  []string
	Data    []user.Human
}

func InitHuman(s *storage.Storage) (Human, error) {
	data, err := s.HumanSelect()
	if err != nil {
		return Human{}, err
	}
	return Human{
		Content: "Информация о человеке",
		Data:    data,
		Fields:  []string{"id", "surname", "name", "mid_name", "passport_series", "passport_num", "birthday", "reg_address"},
	}, nil
}

type RegistrationAddress struct {
	Content string
	Fields  []string
	Data    []user.RegistrationAddress
}

func InitRegistrationAddress(s *storage.Storage) (RegistrationAddress, error) {
	data, err := s.RegistrationAddressSelect()
	if err != nil {
		return RegistrationAddress{}, err
	}
	return RegistrationAddress{
		Content: "Адрес регистрации",
		Data:    data,
		Fields:  []string{"id", "city", "street", "house num", "flat", "country"},
	}, nil
}

type Rent struct {
	Content string
	Fields  []string
	Data    []user.Rent
}

func InitRent(s *storage.Storage) (Rent, error) {
	data, err := s.RentSelect()
	if err != nil {
		return Rent{}, err
	}
	return Rent{
		Content: "Аренда",
		Data:    data,
		Fields:  []string{"id", "transport id", "user id", "cost per hour", "begin date", "end date", "city"},
	}, nil
}

type Transport struct {
	Content string
	Fields  []string
	Data    []transport.Transport
}

func InitTransport(s *storage.Storage) (Transport, error) {
	data, err := s.TransportSelect()
	if err != nil {
		return Transport{}, err
	}
	return Transport{
		Content: "Транспорт",
		Data:    data,
		Fields:  []string{"id", "general info id", "transport info id", "free", "state number", "date add"},
	}, nil
}

type TransportDamages struct {
	Content string
	Fields  []string
	Data    []transport.Damages
}

func InitTransportDamages(s *storage.Storage) (TransportDamages, error) {
	data, err := s.TransportDamagesSelect()
	if err != nil {
		return TransportDamages{}, err
	}
	return TransportDamages{
		Content: "Урон Машин",
		Data:    data,
		Fields:  []string{"id", "transport id", "damage id"},
	}, nil
}

type TransportInfo struct {
	Content string
	Fields  []string
	Data    []transport.Info
}

func InitTransportInfo(s *storage.Storage) (TransportInfo, error) {
	data, err := s.TransportInfoSelect()
	if err != nil {
		return TransportInfo{}, err
	}
	return TransportInfo{
		Content: "Информация о машине",
		Data:    data,
		Fields:  []string{"id", "brand", "release year", "model", "license level"},
	}, nil
}

type User struct {
	Content string
	Fields  []string
	Data    []user.User
}

func InitUser(s *storage.Storage) (User, error) {
	data, err := s.UserSelect()
	if err != nil {
		return User{}, err
	}
	return User{
		Content: "Пользователь",
		Data:    data,
		Fields:  []string{"id", "Nick", "Email", "bank data id", "phone num", "reg date", "human id"},
	}, nil
}

type UserFines struct {
	Content string
	Fields  []string
	Data    []bank.UserFines
}

func InitUserFines(s *storage.Storage) (UserFines, error) {
	data, err := s.UserFinesSelect()
	if err != nil {
		return UserFines{}, err
	}
	return UserFines{
		Content: "пользователь-штраф",
		Data:    data,
		Fields:  []string{"id", "user id", "fine id", "date"},
	}, nil
}

type UserLicense struct {
	Content string
	Fields  []string
	Data    []user.License
}

func InitUserLicense(s *storage.Storage) (UserLicense, error) {
	data, err := s.UserLicenceSelect()
	if err != nil {
		return UserLicense{}, err
	}
	return UserLicense{
		Content: "пользователь-лицензия",
		Data:    data,
		Fields:  []string{"id", "user id", "license id"},
	}, nil
}
