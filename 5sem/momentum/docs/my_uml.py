import os
import subprocess

uml_path = os.path.join(os.getcwd(), 'docs/diagrams')
os.makedirs(uml_path, exist_ok=True)
plantuml_jar_path = './docs/plantuml-1.2024.7.jar'  # Укажите здесь путь к plantuml.jar


def create_uml_file(file_name, uml_code):
    file_path = os.path.join(uml_path, file_name)
    with open(file_path, 'w') as f:
        f.write(uml_code)
    return file_path


def generate_diagram(file_path):
    try:
        print(f"Запуск PlantUML для файла: {file_path}")
        subprocess.run(["java", "-jar", plantuml_jar_path, file_path], check=True)
        print(f"Диаграмма успешно сгенерирована: {file_path}")
    except subprocess.CalledProcessError as e:
        print(f"Ошибка при генерации диаграммы {file_path}: {e}")
    except FileNotFoundError:
        print("Не удалось найти Java. Убедитесь, что Java установлена и доступна.")



sequence_diagram_code = """
@startuml
skinparam DefaultFontSize 25
skinparam ParticipantFontSize 16
skinparam ActorFontSize 25
skinparam ParticipantPadding 20
skinparam BoxPadding 25

actor User
participant "Клиент" as Browser
participant "Сервер" as Backend
participant "База Данных" as Database

User -> Browser: Запрос "Подписаться"
activate Browser
Browser -> Backend: Отправка запроса на подписку
activate Backend
Backend -> Database: Проверить, вошел ли пользователь
activate Database
Database --> Backend: Ответ (Пользователь не аутентифицирован)
deactivate Database


Backend -> Browser: Перенаправление на страницу входа
deactivate Backend
Browser -> User: Пожалуйста войдите в аккаунт
deactivate Browser

User -> Browser: Ввод учетных данных
activate Browser
Browser -> Backend: Отправка учетных данных
activate Backend
Backend -> Database: Проверка учетных данных
activate Database
Database --> Backend: Учетные данные верны
deactivate Database
Backend -> Browser: Перенаправление обратно в профиль
deactivate Backend

Browser -> Backend: Повторный запрос на подписку
activate Backend
Backend -> Database: Проверить наличие подписки
activate Database
Database --> Backend: Подписка отсутствует
deactivate Database
Backend -> Database: Создание подписки
activate Database
Database --> Backend: Подписка создана
deactivate Database
Backend -> Browser: Уведомление об успешной подписке
deactivate Backend
Browser -> User: Вы подписались на пользователя
deactivate Browser
@enduml
"""
sequence_diagram_path = create_uml_file("sequence_diagram.puml", sequence_diagram_code)
generate_diagram(sequence_diagram_path)