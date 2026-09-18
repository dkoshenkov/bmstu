BINARY_NAME=server

S=sudo

DC = sudo docker compose
STOP = sudo docker stop
LOG = sudo docker logs

RM = sudo docker rm
RMI = sudo docker rmi

NGINX_app = nginx-app
NGINX_image = nginx-app-image
GUNI_app=g-app
GUNI_image=g-app-image
APP = app
APP_image =app-image


default: help

help:
	@echo 'Использование: sudo make [target]'
	@echo ''
	@echo 'Сборка и установка:'
	@echo '   build - компиляция образа django'
	@echo '   install - установка зависимостей'
	@echo '   save - сохранение зависимостей в requirements.txt'
	@echo ''
	@echo 'Запуск и управление контейнерами:'
	@echo '   run - поднимает все контейнеры'
	@echo '   r-nginx - перезапускает контейнер nginx'
	@echo '   r-guni - перезапускает контейнер gunicorn'
	@echo '   r-app - перезапускает контейнер приложения'
	@echo '   r-all - перезапускает все контейнеры'
	@echo ''
	@echo 'Просмотр логов и списка контейнеров:'
	@echo '   log-nginx - показывает лог контейнера nginx'
	@echo '   log-guni - показывает лог контейнера gunicorn'
	@echo '   log-app - показывает лог контейнера приложения'
	@echo '   ps - показывает список работающих контейнеров'
	@echo '   psa - показывает список всех контейнеров'
	@echo ''
	@echo 'Остановка и удаление контейнеров:'
	@echo '   down - останавливает и удаляет все контейнеры'
	@echo '   rm-all - удаляет все контейнеры'


doc:
	python3 ./docs/my_uml.py

fill-100:
	python3 manage.py fill_db 100

install:
	pip install -r requirements.txt

run:
	python3 manage.py runserver localhost:8888

guni-run:
	gunicorn -b 0.0.0.0:9009 momentum.wsgi:application &

rm-all:
	$(STOP) $(NGINX_app) $(APP) && $(RM) $(NGINX_app)

save:
	pip freeze > requirements.txt

ps:
	docker ps

psa:
	$(S) docker ps -a

build:
	docker build -t app-image .

#run:
#	$(DC) up -d

r-nginx:
	 $(STOP) $(NGINX_app) &&  $(RM)  $(NGINX_app) &&  $(RMI) $(NGINX_image) &&  $(DC) up -d

r-guni:
	 $(STOP) $(GUNI_app) && $(RM) $(GUNI_app) && $(RMI) $(GUNI_image) &&  $(DC) up -d

r-app:
	$(STOP) $(APP) && $(RM) $(APP) && $(RMI) $(APP_image) &&  $(DC) up -d

r-all:
	make r-nginx
	make r-guni
	make r-app

log-nginx:
	$(LOG) $(NGINX_app)

log-guni:
	$(LOG) $(GUNI_app)

log-app:
	$(LOG) $(APP)

down:
	$(DC) down

.PHONY: clean build
