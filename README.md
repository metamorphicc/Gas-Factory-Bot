# Telegram Document Assistant MVP

Демо-MVP Telegram-бота для работы одного владельца с клиентскими папками и Word-шаблонами.

## Модель хранения

Бот хранит файлы там, где он запущен.

Локально:

```env
STORAGE_ROOT=data
```

Файлы будут лежать в:

```text
data/clients/<client_name>/
```

На VPS:

```env
STORAGE_ROOT=/var/docbot/data
```

Файлы будут лежать в:

```text
/var/docbot/data/clients/<client_name>/
```

То есть код один и тот же: локально хранит локально, на VPS хранит на VPS.

## Возможности

- один владелец бота: первый пользователь, который запустил демо, получает доступ
- добавить клиента и создать папку `clients/<client_name>/`
- загрузить файл клиента: `pdf`, `doc`, `docx`, `jpg`, `png`
- посмотреть клиентские папки и скачать файл обратно
- заполнить `.docx` шаблон по найденным полям `{{field}}` одним сообщением
- сохранить готовый документ в папку клиента и отправить его в чат

## Запуск локально

```bash
pip install -r requirements.txt
```

Создайте `.env` рядом с `main.py`:

```env
BOT_TOKEN=your_telegram_bot_token
STORAGE_ROOT=data
```

Запустите:

```bash
python main.py
```

## Запуск на VPS

```bash
sudo apt update
sudo apt install -y python3 python3-venv python3-pip git
git clone <repo_url> docbot
cd docbot
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
mkdir -p /var/docbot/data
```

Создайте `.env`:

```env
BOT_TOKEN=your_telegram_bot_token
STORAGE_ROOT=/var/docbot/data
```

Тестовый запуск:

```bash
python main.py
```

## Systemd для VPS

```ini
[Unit]
Description=Telegram document assistant bot
After=network.target

[Service]
WorkingDirectory=/home/ubuntu/docbot
ExecStart=/home/ubuntu/docbot/.venv/bin/python /home/ubuntu/docbot/main.py
Restart=always
RestartSec=5
User=ubuntu

[Install]
WantedBy=multi-user.target
```

Команды:

```bash
sudo systemctl daemon-reload
sudo systemctl enable docbot
sudo systemctl start docbot
sudo systemctl status docbot
```

## Заполнение документа

Бот просит все поля одним сообщением. Можно скопировать форму из чата и заполнить так:

```text
ФИО: Иванов Иван Иванович
Паспорт: 1234 567890
ИНН: 123456789012
Дата: 28.08.2026
```

После этого бот покажет карточку подтверждения. Паспорт и ИНН в карточке маскируются.
Если данные неверные, нажмите `Изменить данные` - бот вернет заполненную форму на шаг назад, и ее можно будет отправить заново.

Во время сценариев бот удаляет служебные сообщения текущего шага после завершения действия, возврата в меню или перехода к редактированию. В чате остаются финальные результаты: подтверждение, путь к файлу или готовый документ.

## Папки

- `templates/` - Word-шаблоны `.docx`
- `<STORAGE_ROOT>/clients/` - клиентские папки и документы

В проекте есть демо-шаблон `templates/demo_client_form.docx` с полями:

- `{{full_name}}`
- `{{passport}}`
- `{{inn}}`
- `{{date}}`

## Безопасность

Это MVP demo для клиента, не production IAM. Доступ ограничен первым Telegram-пользователем, который запустил бота; его id хранится в `<STORAGE_ROOT>/owner.txt`. Файлы ограничены по типу и размеру, имена папок и файлов очищаются перед сохранением.
