import asyncio
import logging
import os
import re
import zipfile
from datetime import datetime
from pathlib import Path
from typing import Any

from aiogram import Bot, Dispatcher, F, Router
from aiogram.exceptions import TelegramAPIError
from aiogram.filters import CommandStart
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.fsm.storage.memory import MemoryStorage
from aiogram.types import (
    CallbackQuery,
    FSInputFile,
    InlineKeyboardButton,
    InlineKeyboardMarkup,
    KeyboardButton,
    Message,
    ReplyKeyboardMarkup,
)
from docx import Document
from dotenv import load_dotenv
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill


BASE_DIR = Path(__file__).resolve().parent
load_dotenv()


def resolve_storage_root(raw_path: str) -> Path:
    path = Path(raw_path).expanduser()
    if not path.is_absolute():
        path = BASE_DIR / path
    return path.resolve()


DATA_DIR = resolve_storage_root(os.getenv("STORAGE_ROOT", "data"))
CLIENTS_DIR = DATA_DIR / "clients"
TEMPLATES_DIR = BASE_DIR / "templates"
CLIENTS_EXCEL_PATH = DATA_DIR / "clients.xlsx"

ALLOWED_EXTENSIONS = {".pdf", ".doc", ".docx", ".jpg", ".jpeg", ".png"}
MAX_FILE_SIZE = 20 * 1024 * 1024
PLACEHOLDER_RE = re.compile(r"{{\s*([a-zA-Z_][a-zA-Z0-9_]*)\s*}}")
FLOW_MESSAGES_KEY = "flow_message_ids"
SCREEN_MESSAGE_IDS: dict[int, list[int]] = {}
CLIENTS_SHEET_NAME = "clients"
CLIENTS_HEADERS = [
    "created_at",
    "updated_at",
    "client_name",
    "folder",
    "files_count",
    "last_file",
    "full_name",
    "passport",
    "inn",
    "date",
    "last_template",
    "last_document",
]

FIELD_LABELS = {
    "full_name": "ФИО",
    "passport": "Паспорт",
    "inn": "ИНН",
    "date": "Дата",
}

MAIN_BUTTONS = {
    "Добавить клиента",
    "Добавить файл",
    "Получить файл",
    "Заполнить документ",
    "Мои клиенты",
}
NAV_BUTTONS = {"Меню", "В меню"}


class BotStates(StatesGroup):
    waiting_client_name = State()
    add_file_choose_client = State()
    add_file_waiting_file = State()
    get_file_choose_client = State()
    get_file_choose_file = State()
    fill_template_choose_template = State()
    fill_template_choose_client = State()
    fill_template_ask_field = State()
    fill_template_confirm = State()


router = Router()


def ensure_storage() -> None:
    CLIENTS_DIR.mkdir(parents=True, exist_ok=True)
    TEMPLATES_DIR.mkdir(parents=True, exist_ok=True)
    ensure_clients_workbook()


def now_text() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def style_clients_sheet(workbook: Workbook) -> None:
    worksheet = workbook[CLIENTS_SHEET_NAME]
    header_fill = PatternFill("solid", fgColor="D9EAF7")
    for cell in worksheet[1]:
        cell.font = Font(bold=True)
        cell.fill = header_fill
    worksheet.freeze_panes = "A2"
    widths = {
        "A": 19,
        "B": 19,
        "C": 24,
        "D": 38,
        "E": 12,
        "F": 32,
        "G": 28,
        "H": 18,
        "I": 18,
        "J": 14,
        "K": 28,
        "L": 38,
    }
    for column, width in widths.items():
        worksheet.column_dimensions[column].width = width


def ensure_clients_workbook() -> None:
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    if not CLIENTS_EXCEL_PATH.exists():
        workbook = Workbook()
        worksheet = workbook.active
        worksheet.title = CLIENTS_SHEET_NAME
        worksheet.append(CLIENTS_HEADERS)
        style_clients_sheet(workbook)
        workbook.save(CLIENTS_EXCEL_PATH)
        return

    workbook = load_workbook(CLIENTS_EXCEL_PATH)
    if CLIENTS_SHEET_NAME not in workbook.sheetnames:
        worksheet = workbook.create_sheet(CLIENTS_SHEET_NAME)
        worksheet.append(CLIENTS_HEADERS)
        style_clients_sheet(workbook)
        workbook.save(CLIENTS_EXCEL_PATH)
        return

    worksheet = workbook[CLIENTS_SHEET_NAME]
    existing_headers = [cell.value for cell in worksheet[1]]
    changed = False
    for header in CLIENTS_HEADERS:
        if header not in existing_headers:
            worksheet.cell(row=1, column=len(existing_headers) + 1, value=header)
            existing_headers.append(header)
            changed = True

    if changed:
        style_clients_sheet(workbook)
        workbook.save(CLIENTS_EXCEL_PATH)


def clients_header_map(worksheet: Any) -> dict[str, int]:
    return {cell.value: cell.column for cell in worksheet[1] if cell.value}


def find_client_row(worksheet: Any, headers: dict[str, int], client_name: str) -> int | None:
    client_column = headers["client_name"]
    for row in range(2, worksheet.max_row + 1):
        if worksheet.cell(row=row, column=client_column).value == client_name:
            return row
    return None


def upsert_client_record(client_name: str, updates: dict[str, Any] | None = None) -> None:
    ensure_clients_workbook()
    workbook = load_workbook(CLIENTS_EXCEL_PATH)
    worksheet = workbook[CLIENTS_SHEET_NAME]
    headers = clients_header_map(worksheet)
    row = find_client_row(worksheet, headers, client_name)
    current_time = now_text()

    if row is None:
        row = worksheet.max_row + 1
        worksheet.cell(row=row, column=headers["created_at"], value=current_time)
        worksheet.cell(row=row, column=headers["client_name"], value=client_name)

    worksheet.cell(row=row, column=headers["updated_at"], value=current_time)
    for key, value in (updates or {}).items():
        if key in headers:
            worksheet.cell(row=row, column=headers[key], value=value)

    style_clients_sheet(workbook)
    workbook.save(CLIENTS_EXCEL_PATH)


def main_menu_keyboard() -> ReplyKeyboardMarkup:
    return ReplyKeyboardMarkup(
        keyboard=[
            [KeyboardButton(text="Добавить клиента"), KeyboardButton(text="Добавить файл")],
            [KeyboardButton(text="Получить файл"), KeyboardButton(text="Заполнить документ")],
            [KeyboardButton(text="Мои клиенты")],
        ],
        resize_keyboard=True,
        input_field_placeholder="Выберите действие",
    )


def menu_inline_keyboard() -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup(
        inline_keyboard=[[InlineKeyboardButton(text="В меню", callback_data="menu")]]
    )


def upload_done_keyboard() -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup(
        inline_keyboard=[
            [InlineKeyboardButton(text="Оставить как есть", callback_data="add_file:done")],
            [InlineKeyboardButton(text="В меню", callback_data="menu")],
        ]
    )


def confirmation_keyboard() -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup(
        inline_keyboard=[
            [InlineKeyboardButton(text="Все верно", callback_data="fill:ok")],
            [InlineKeyboardButton(text="Изменить данные", callback_data="fill:edit")],
            [InlineKeyboardButton(text="В меню", callback_data="menu")],
        ]
    )


def rows_keyboard(prefix: str, items: list[str]) -> InlineKeyboardMarkup:
    rows = [
        [InlineKeyboardButton(text=item, callback_data=f"{prefix}:{index}")]
        for index, item in enumerate(items)
    ]
    rows.append([InlineKeyboardButton(text="В меню", callback_data="menu")])
    return InlineKeyboardMarkup(inline_keyboard=rows)


async def remember_flow_message(state: FSMContext, message: Message) -> None:
    data = await state.get_data()
    messages = data.get(FLOW_MESSAGES_KEY, [])
    item = {"chat_id": message.chat.id, "message_id": message.message_id}
    if item not in messages:
        messages.append(item)
        await state.update_data(**{FLOW_MESSAGES_KEY: messages})


def remember_screen_message(message: Message) -> None:
    message_ids = SCREEN_MESSAGE_IDS.setdefault(message.chat.id, [])
    if message.message_id not in message_ids:
        message_ids.append(message.message_id)


async def answer_flow(message: Message, state: FSMContext, text: str, **kwargs: Any) -> Message:
    sent = await message.answer(text, **kwargs)
    remember_screen_message(sent)
    await remember_flow_message(state, sent)
    return sent


async def answer_final(message: Message, state: FSMContext, text: str, **kwargs: Any) -> Message:
    await state.clear()
    return await message.answer(text, **kwargs)


async def delete_message_safely(bot: Bot, chat_id: int, message_id: int) -> None:
    try:
        await bot.delete_message(chat_id=chat_id, message_id=message_id)
    except TelegramAPIError:
        pass


async def clear_screen_messages(bot: Bot, chat_id: int) -> None:
    message_ids = SCREEN_MESSAGE_IDS.pop(chat_id, [])
    for message_id in reversed(message_ids):
        await delete_message_safely(bot, chat_id, message_id)


async def clear_flow_messages(bot: Bot, state: FSMContext, chat_id: int | None = None) -> None:
    data = await state.get_data()
    messages = data.get(FLOW_MESSAGES_KEY, [])
    seen: set[tuple[int, int]] = set()
    for item in reversed(messages):
        key = (item["chat_id"], item["message_id"])
        if key not in seen:
            seen.add(key)
            await delete_message_safely(bot, item["chat_id"], item["message_id"])

    chat_ids = {item["chat_id"] for item in messages}
    if chat_id is not None:
        chat_ids.add(chat_id)
    for current_chat_id in chat_ids:
        await clear_screen_messages(bot, current_chat_id)

    await state.update_data(**{FLOW_MESSAGES_KEY: []})


def sanitize_name(raw_name: str) -> str:
    name = raw_name.strip()
    name = re.sub(r"[^\w\s-]", "_", name, flags=re.UNICODE)
    name = re.sub(r"\s+", "_", name)
    name = re.sub(r"_+", "_", name)
    name = name.strip("._-")[:60]
    return name or "client"


def safe_child_path(root: Path, name: str) -> Path:
    candidate = (root / name).resolve()
    root_resolved = root.resolve()
    if root_resolved != candidate and root_resolved not in candidate.parents:
        raise ValueError("Path escapes storage root")
    return candidate


def display_path(path: Path) -> str:
    resolved = path.resolve()
    for root in (BASE_DIR.resolve(), DATA_DIR.resolve()):
        try:
            return resolved.relative_to(root).as_posix()
        except ValueError:
            continue
    return resolved.as_posix()


def client_names() -> list[str]:
    ensure_storage()
    return sorted(path.name for path in CLIENTS_DIR.iterdir() if path.is_dir())


def template_names() -> list[str]:
    ensure_storage()
    return sorted(
        path.name
        for path in TEMPLATES_DIR.iterdir()
        if path.is_file() and path.suffix.lower() == ".docx" and not path.name.startswith("~$")
    )


def file_names(client_name: str) -> list[str]:
    folder = safe_child_path(CLIENTS_DIR, client_name)
    return sorted(path.name for path in folder.iterdir() if path.is_file())


def get_template_fields(template_path: Path) -> list[str]:
    found: set[str] = set()
    with zipfile.ZipFile(template_path) as docx_zip:
        for member in docx_zip.namelist():
            if member.startswith("word/") and member.endswith(".xml"):
                text = docx_zip.read(member).decode("utf-8", errors="ignore")
                found.update(PLACEHOLDER_RE.findall(text))

    ordered = [field for field in FIELD_LABELS if field in found]
    ordered.extend(sorted(found - set(ordered)))
    return ordered


def replace_in_paragraph(paragraph: Any, values: dict[str, str]) -> None:
    original = paragraph.text
    updated = original
    for field, value in values.items():
        updated = re.sub(r"{{\s*" + re.escape(field) + r"\s*}}", value, updated)
    if updated != original:
        paragraph.text = updated


def fill_docx(template_path: Path, output_path: Path, values: dict[str, str]) -> None:
    document = Document(template_path)
    for paragraph in document.paragraphs:
        replace_in_paragraph(paragraph, values)
    for table in document.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    replace_in_paragraph(paragraph, values)
    document.save(output_path)


def unique_path(folder: Path, filename: str) -> Path:
    sanitized = sanitize_filename(filename)
    target = safe_child_path(folder, sanitized)
    if not target.exists():
        return target

    stem = target.stem
    suffix = target.suffix
    for number in range(1, 1000):
        candidate = safe_child_path(folder, f"{stem}_{number}{suffix}")
        if not candidate.exists():
            return candidate
    raise RuntimeError("Too many files with the same name")


def sanitize_filename(raw_filename: str) -> str:
    path = Path(raw_filename)
    stem = sanitize_name(path.stem)
    suffix = path.suffix.lower()
    return f"{stem}{suffix}" if suffix else stem


def rename_target_path(source_path: Path, requested_name: str) -> Path:
    clean_stem = sanitize_name(Path(requested_name.strip()).stem)
    filename = f"{clean_stem}{source_path.suffix.lower()}"
    if filename == source_path.name:
        return source_path
    return unique_path(source_path.parent, filename)


def build_upload_rename_prompt(uploaded_files: list[dict[str, str]]) -> str:
    count = len(uploaded_files)
    if count == 1:
        return (
            "Файл сохранен.\n\n"
            "Если хотите присвоить ему нормальное имя, напишите его одним сообщением.\n"
            "Расширение файла сохранится автоматически."
        )

    examples = "\n".join("название файла" for _ in uploaded_files)
    return (
        f"Распознаны файлы: {count} шт.\n\n"
        "Если хотите присвоить им нормальные имена, напишите их в формате по одному имени на строку:\n\n"
        f"{examples}\n\n"
        "Расширения файлов сохранятся автоматически."
    )


def uploaded_files_summary(uploaded_files: list[dict[str, str]]) -> str:
    lines = [f"Файлы сохранены: {len(uploaded_files)} шт."]
    for file_info in uploaded_files:
        lines.append(f"- {file_info['display_path']}")
    return "\n".join(lines)


def rename_uploaded_files(uploaded_files: list[dict[str, str]], names: list[str]) -> list[dict[str, str]]:
    renamed: list[dict[str, str]] = []
    for file_info, requested_name in zip(uploaded_files, names):
        source_path = Path(file_info["path"])
        if not source_path.exists():
            renamed.append(file_info)
            continue

        target_path = rename_target_path(source_path, requested_name)
        if target_path != source_path:
            source_path.replace(target_path)

        renamed.append(
            {
                "path": str(target_path),
                "display_path": display_path(target_path),
            }
        )
    return renamed


def update_client_files_record(client_name: str, last_file_path: Path | None = None) -> None:
    client_folder = safe_child_path(CLIENTS_DIR, client_name)
    files = file_names(client_name)
    updates: dict[str, Any] = {
        "folder": display_path(client_folder),
        "files_count": len(files),
    }
    if last_file_path is not None:
        updates["last_file"] = display_path(last_file_path)
    elif files:
        updates["last_file"] = display_path(safe_child_path(client_folder, files[-1]))
    upsert_client_record(client_name, updates)


def mask_value(field: str, value: str) -> str:
    if field not in {"passport", "inn"}:
        return value
    digits = re.sub(r"\D", "", value)
    tail = digits[-4:] if len(digits) >= 4 else value[-2:]
    return f"***{tail}"


def field_label(field: str) -> str:
    return FIELD_LABELS.get(field, field)


def normalize_label(label: str) -> str:
    return re.sub(r"[\W_]+", "", label.lower(), flags=re.UNICODE)


def build_fields_prompt(fields: list[str], values: dict[str, str] | None = None) -> str:
    lines = [
        "Заполните данные одним сообщением.",
        "Скопируйте форму ниже и впишите значения:",
        "",
    ]
    lines.extend(f"{field_label(field)}: {(values or {}).get(field, '')}" for field in fields)
    return "\n".join(lines)


def parse_document_values(text: str, fields: list[str]) -> tuple[dict[str, str], list[str]]:
    label_to_field: dict[str, str] = {}
    for field in fields:
        label_to_field[normalize_label(field)] = field
        label_to_field[normalize_label(field_label(field))] = field

    values: dict[str, str] = {}
    positional_lines: list[str] = []
    for raw_line in text.splitlines():
        line = raw_line.strip()
        if not line:
            continue

        if ":" in line:
            label, value = line.split(":", 1)
            field = label_to_field.get(normalize_label(label))
            if field:
                values[field] = value.strip()
                continue

        positional_lines.append(line)

    remaining_fields = [field for field in fields if not values.get(field)]
    if len(positional_lines) == len(remaining_fields):
        for field, value in zip(remaining_fields, positional_lines):
            values[field] = value.strip()

    missing = [field for field in fields if not values.get(field)]
    return values, missing


async def show_main_menu(message: Message, state: FSMContext) -> None:
    await state.clear()
    await answer_flow(message, state, "Главное меню", reply_markup=main_menu_keyboard())


async def show_client_list(message: Message, state: FSMContext) -> None:
    clients = client_names()
    if not clients:
        await answer_flow(message, state, "Главное меню\n\nКлиентов пока нет.", reply_markup=main_menu_keyboard())
        return

    text = "Главное меню\n\nМои клиенты:\n" + "\n".join(f"- {client}" for client in clients)
    await answer_flow(message, state, text, reply_markup=main_menu_keyboard())


async def show_clients_for_flow(message: Message, state: FSMContext, flow: str) -> None:
    clients = client_names()
    if not clients:
        await answer_flow(
            message,
            state,
            "Пока нет клиентов. Сначала добавьте клиента.",
            reply_markup=menu_inline_keyboard(),
        )
        return

    await state.update_data(clients=clients)
    if flow == "add":
        await state.set_state(BotStates.add_file_choose_client)
        title = "Выберите клиента, куда сохранить файл:"
        prefix = "add_client"
    elif flow == "get":
        await state.set_state(BotStates.get_file_choose_client)
        title = "Выберите клиента:"
        prefix = "get_client"
    else:
        await state.set_state(BotStates.fill_template_choose_client)
        title = "Выберите клиента для готового документа:"
        prefix = "fill_client"

    await answer_flow(message, state, title, reply_markup=rows_keyboard(prefix, clients))


@router.message(CommandStart())
async def start(message: Message, state: FSMContext, bot: Bot) -> None:
    await clear_flow_messages(bot, state, message.chat.id)
    await delete_message_safely(bot, message.chat.id, message.message_id)
    await show_main_menu(message, state)


@router.callback_query(F.data == "menu")
async def callback_menu(callback: CallbackQuery, state: FSMContext, bot: Bot) -> None:
    chat_id = callback.message.chat.id if callback.message else None
    await clear_flow_messages(bot, state, chat_id)
    await state.clear()
    await callback.answer()
    if callback.message:
        await answer_flow(callback.message, state, "Главное меню", reply_markup=main_menu_keyboard())


@router.message(F.text.in_(MAIN_BUTTONS | NAV_BUTTONS))
async def main_menu_buttons(message: Message, state: FSMContext, bot: Bot) -> None:

    await clear_flow_messages(bot, state, message.chat.id)
    await delete_message_safely(bot, message.chat.id, message.message_id)
    await state.clear()
    if message.text in NAV_BUTTONS:
        await answer_flow(message, state, "Главное меню", reply_markup=main_menu_keyboard())
    elif message.text == "Добавить клиента":
        await state.set_state(BotStates.waiting_client_name)
        await answer_flow(
            message,
            state,
            "Введите имя или короткий ID клиента:",
            reply_markup=menu_inline_keyboard(),
        )
    elif message.text == "Добавить файл":
        await show_clients_for_flow(message, state, "add")
    elif message.text == "Получить файл":
        await show_clients_for_flow(message, state, "get")
    elif message.text == "Заполнить документ":
        await fill_template_start(message, state)
    elif message.text == "Мои клиенты":
        await show_client_list(message, state)


@router.message(F.text == "Добавить клиента")
async def add_client_start(message: Message, state: FSMContext) -> None:
    await state.set_state(BotStates.waiting_client_name)
    await answer_flow(
        message,
        state,
        "Введите имя или короткий ID клиента:",
        reply_markup=menu_inline_keyboard(),
    )


@router.message(BotStates.waiting_client_name)
async def add_client_finish(message: Message, state: FSMContext, bot: Bot) -> None:
    await remember_flow_message(state, message)
    if not message.text:
        await answer_flow(message, state, "Пришлите имя текстом.", reply_markup=menu_inline_keyboard())
        return

    client_name = sanitize_name(message.text)
    client_path = safe_child_path(CLIENTS_DIR, client_name)
    await clear_flow_messages(bot, state, message.chat.id)
    if client_path.exists():
        update_client_files_record(client_name)
        await message.answer(
            f"Главное меню\n\nКлиент уже есть: {display_path(client_path)}",
            reply_markup=main_menu_keyboard(),
        )
    else:
        client_path.mkdir(parents=True, exist_ok=False)
        upsert_client_record(
            client_name,
            {
                "folder": display_path(client_path),
                "files_count": 0,
            },
        )
        await message.answer(
            f"Главное меню\n\nКлиент добавлен: {display_path(client_path)}",
            reply_markup=main_menu_keyboard(),
        )
    await state.clear()


@router.message(F.text == "Мои клиенты")
async def my_clients(message: Message, state: FSMContext) -> None:
    await state.clear()
    await show_client_list(message, state)


@router.message(F.text == "Добавить файл")
async def add_file_start(message: Message, state: FSMContext) -> None:
    await show_clients_for_flow(message, state, "add")


@router.callback_query(BotStates.add_file_choose_client, F.data.startswith("add_client:"))
async def add_file_choose_client(callback: CallbackQuery, state: FSMContext) -> None:
    data = await state.get_data()
    clients = data.get("clients", [])
    index = int(callback.data.split(":", 1)[1])
    if index >= len(clients):
        await callback.answer("Клиент не найден", show_alert=True)
        return

    await state.update_data(selected_client=clients[index], uploaded_files=[])
    await state.set_state(BotStates.add_file_waiting_file)
    await callback.answer()
    if callback.message:
        await answer_flow(
            callback.message,
            state,
            "Пришлите один или несколько файлов: pdf, doc, docx, jpg или png. Максимум 20 МБ на файл.",
            reply_markup=menu_inline_keyboard(),
        )


@router.message(BotStates.add_file_waiting_file)
async def add_file_receive(message: Message, state: FSMContext, bot: Bot) -> None:
    await remember_flow_message(state, message)

    data = await state.get_data()
    uploaded_files = data.get("uploaded_files", [])

    if message.text and uploaded_files:
        if not message.text.strip():
            await answer_flow(message, state, "Пришлите непустое имя файла.", reply_markup=upload_done_keyboard())
            return

        names = [line.strip() for line in message.text.splitlines() if line.strip()]
        if len(uploaded_files) == 1:
            names = [message.text.strip()]
        elif len(names) != len(uploaded_files):
            await answer_flow(
                message,
                state,
                (
                    f"Нужно прислать {len(uploaded_files)} имен, по одному на строку.\n\n"
                    + build_upload_rename_prompt(uploaded_files)
                ),
                reply_markup=upload_done_keyboard(),
            )
            return

        renamed_files = rename_uploaded_files(uploaded_files, names)
        if renamed_files:
            update_client_files_record(data["selected_client"], Path(renamed_files[-1]["path"]))
        await state.update_data(uploaded_files=renamed_files)
        await clear_flow_messages(bot, state, message.chat.id)
        await state.clear()
        await message.answer(
            "Главное меню\n\n"
            + uploaded_files_summary(renamed_files),
            reply_markup=main_menu_keyboard(),
        )
        return

    telegram_file_id: str | None = None
    original_name: str | None = None
    file_size = 0

    if message.document:
        telegram_file_id = message.document.file_id
        original_name = message.document.file_name or "file"
        file_size = message.document.file_size or 0
    elif message.photo:
        photo = message.photo[-1]
        telegram_file_id = photo.file_id
        original_name = f"photo_{photo.file_unique_id}.jpg"
        file_size = photo.file_size or 0

    if not telegram_file_id or not original_name:
        await answer_flow(message, state, "Пришлите файл или фото.", reply_markup=menu_inline_keyboard())
        return

    if file_size > MAX_FILE_SIZE:
        await answer_flow(message, state, "Файл слишком большой. Максимум 20 МБ.", reply_markup=menu_inline_keyboard())
        return

    extension = Path(original_name).suffix.lower()
    if extension not in ALLOWED_EXTENSIONS:
        await answer_flow(message, state, "Можно загрузить только pdf, doc, docx, jpg или png.", reply_markup=menu_inline_keyboard())
        return

    data = await state.get_data()
    client_name = data["selected_client"]
    client_folder = safe_child_path(CLIENTS_DIR, client_name)
    target_path = unique_path(client_folder, original_name)

    telegram_file = await bot.get_file(telegram_file_id)
    remote_size = telegram_file.file_size or file_size
    if remote_size > MAX_FILE_SIZE:
        await answer_flow(message, state, "Файл слишком большой. Максимум 20 МБ.", reply_markup=menu_inline_keyboard())
        return

    await bot.download_file(telegram_file.file_path, destination=target_path)
    update_client_files_record(client_name, target_path)
    saved_files = data.get("uploaded_files", [])
    saved_files.append(
        {
            "path": str(target_path),
            "display_path": display_path(target_path),
        }
    )
    await state.update_data(uploaded_files=saved_files)
    await clear_screen_messages(bot, message.chat.id)
    await answer_flow(
        message,
        state,
        build_upload_rename_prompt(saved_files),
        reply_markup=upload_done_keyboard(),
    )


@router.callback_query(BotStates.add_file_waiting_file, F.data == "add_file:done")
async def add_file_finish_without_rename(callback: CallbackQuery, state: FSMContext, bot: Bot) -> None:
    data = await state.get_data()
    uploaded_files = data.get("uploaded_files", [])
    if not uploaded_files:
        await callback.answer("Сначала пришлите файл", show_alert=True)
        return

    chat_id = callback.message.chat.id if callback.message else None
    update_client_files_record(data["selected_client"], Path(uploaded_files[-1]["path"]))
    await clear_flow_messages(bot, state, chat_id)
    await state.clear()
    await callback.answer()
    if callback.message:
        await callback.message.answer(
            "Главное меню\n\n"
            + uploaded_files_summary(uploaded_files),
            reply_markup=main_menu_keyboard(),
        )


@router.message(F.text == "Получить файл")
async def get_file_start(message: Message, state: FSMContext) -> None:
    await show_clients_for_flow(message, state, "get")


@router.callback_query(BotStates.get_file_choose_client, F.data.startswith("get_client:"))
async def get_file_choose_client(callback: CallbackQuery, state: FSMContext) -> None:
    data = await state.get_data()
    clients = data.get("clients", [])
    index = int(callback.data.split(":", 1)[1])
    if index >= len(clients):
        await callback.answer("Клиент не найден", show_alert=True)
        return

    client_name = clients[index]
    files = file_names(client_name)
    if not files:
        await callback.answer()
        if callback.message:
            await answer_flow(
                callback.message,
                state,
                "В папке клиента пока нет файлов.",
                reply_markup=menu_inline_keyboard(),
            )
        return

    await state.update_data(selected_client=client_name, files=files)
    await state.set_state(BotStates.get_file_choose_file)
    await callback.answer()
    if callback.message:
        await answer_flow(
            callback.message,
            state,
            "Выберите файл:",
            reply_markup=rows_keyboard("get_file", files),
        )


@router.callback_query(BotStates.get_file_choose_file, F.data.startswith("get_file:"))
async def get_file_send(callback: CallbackQuery, state: FSMContext, bot: Bot) -> None:
    data = await state.get_data()
    files = data.get("files", [])
    index = int(callback.data.split(":", 1)[1])
    if index >= len(files):
        await callback.answer("Файл не найден", show_alert=True)
        return

    client_name = data["selected_client"]
    file_path = safe_child_path(safe_child_path(CLIENTS_DIR, client_name), files[index])
    chat_id = callback.message.chat.id if callback.message else None
    await clear_flow_messages(bot, state, chat_id)
    await state.clear()
    await callback.answer()
    if callback.message:
        await callback.message.answer_document(
            FSInputFile(file_path),
            caption=f"Файл: {files[index]}",
            reply_markup=main_menu_keyboard(),
        )


@router.message(F.text == "Заполнить документ")
async def fill_template_start(message: Message, state: FSMContext) -> None:

    templates = template_names()
    if not templates:
        await answer_flow(
            message,
            state,
            "В папке templates нет .docx шаблонов.",
            reply_markup=menu_inline_keyboard(),
        )
        return

    await state.update_data(templates=templates)
    await state.set_state(BotStates.fill_template_choose_template)
    await answer_flow(message, state, "Выберите шаблон:", reply_markup=rows_keyboard("template", templates))


@router.callback_query(BotStates.fill_template_choose_template, F.data.startswith("template:"))
async def fill_template_choose_template(callback: CallbackQuery, state: FSMContext) -> None:
    data = await state.get_data()
    templates = data.get("templates", [])
    index = int(callback.data.split(":", 1)[1])
    if index >= len(templates):
        await callback.answer("Шаблон не найден", show_alert=True)
        return

    template_name = templates[index]
    template_path = safe_child_path(TEMPLATES_DIR, template_name)
    fields = get_template_fields(template_path)
    if not fields:
        await callback.answer()
        if callback.message:
            await answer_flow(
                callback.message,
                state,
                "В шаблоне не найдены поля вида {{field}}.",
                reply_markup=menu_inline_keyboard(),
            )
        return

    await state.update_data(selected_template=template_name, fields=fields, values={})
    await callback.answer()
    if callback.message:
        await show_clients_for_flow(callback.message, state, "fill")


@router.callback_query(BotStates.fill_template_choose_client, F.data.startswith("fill_client:"))
async def fill_template_choose_client(callback: CallbackQuery, state: FSMContext) -> None:
    data = await state.get_data()
    clients = data.get("clients", [])
    index = int(callback.data.split(":", 1)[1])
    if index >= len(clients):
        await callback.answer("Клиент не найден", show_alert=True)
        return

    fields = data["fields"]
    await state.update_data(selected_client=clients[index])
    await state.set_state(BotStates.fill_template_ask_field)
    await callback.answer()
    if callback.message:
        await answer_flow(callback.message, state, build_fields_prompt(fields), reply_markup=menu_inline_keyboard())


@router.message(BotStates.fill_template_ask_field)
async def fill_template_collect_field(message: Message, state: FSMContext) -> None:
    await remember_flow_message(state, message)
    if not message.text:
        await answer_flow(message, state, "Пришлите заполненные данные текстом.", reply_markup=menu_inline_keyboard())
        return

    data = await state.get_data()
    fields = data["fields"]
    values, missing = parse_document_values(message.text, fields)
    if missing:
        missing_labels = ", ".join(field_label(field) for field in missing)
        await answer_flow(
            message,
            state,
            f"Не вижу обязательные поля: {missing_labels}\n\n{build_fields_prompt(fields)}",
            reply_markup=menu_inline_keyboard(),
        )
        return

    await state.update_data(values=values)
    await state.set_state(BotStates.fill_template_confirm)
    lines = [
        "Проверьте данные:",
        f"Клиент: {data['selected_client']}",
        f"Шаблон: {data['selected_template']}",
    ]
    for field in fields:
        label = field_label(field)
        lines.append(f"{label}: {mask_value(field, values[field])}")
    await answer_flow(message, state, "\n".join(lines), reply_markup=confirmation_keyboard())


@router.callback_query(BotStates.fill_template_confirm, F.data == "fill:ok")
async def fill_template_finish(callback: CallbackQuery, state: FSMContext, bot: Bot) -> None:

    data = await state.get_data()
    client_name = data["selected_client"]
    template_name = data["selected_template"]
    values = data["values"]

    client_folder = safe_child_path(CLIENTS_DIR, client_name)
    template_path = safe_child_path(TEMPLATES_DIR, template_name)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_name = f"{Path(template_name).stem}_{timestamp}.docx"
    output_path = unique_path(client_folder, output_name)
    fill_docx(template_path, output_path, values)
    upsert_client_record(
        client_name,
        {
            "folder": display_path(client_folder),
            "files_count": len(file_names(client_name)),
            "full_name": values.get("full_name", ""),
            "passport": values.get("passport", ""),
            "inn": values.get("inn", ""),
            "date": values.get("date", ""),
            "last_template": template_name,
            "last_document": display_path(output_path),
        },
    )

    chat_id = callback.message.chat.id if callback.message else None
    await clear_flow_messages(bot, state, chat_id)
    await state.clear()
    await callback.answer()
    if callback.message:
        relative_path = display_path(output_path)
        await callback.message.answer(
            f"Главное меню\n\nДокумент готов: {relative_path}",
            reply_markup=main_menu_keyboard(),
        )
        await callback.message.answer_document(FSInputFile(output_path), caption="Готовый документ")


@router.callback_query(BotStates.fill_template_confirm, F.data == "fill:edit")
async def fill_template_edit(callback: CallbackQuery, state: FSMContext, bot: Bot) -> None:

    data = await state.get_data()
    fields = data["fields"]
    values = data.get("values", {})
    chat_id = callback.message.chat.id if callback.message else None
    await clear_flow_messages(bot, state, chat_id)
    await state.set_state(BotStates.fill_template_ask_field)
    await callback.answer()
    if callback.message:
        await answer_flow(
            callback.message,
            state,
            "Ок, поправьте данные и пришлите форму заново:\n\n"
            + build_fields_prompt(fields, values),
            reply_markup=menu_inline_keyboard(),
        )


@router.message(F.text.in_({"Меню", "В меню"}))
async def menu_text(message: Message, state: FSMContext) -> None:
    await show_main_menu(message, state)


@router.message()
async def fallback(message: Message, state: FSMContext) -> None:
    await message.answer("Главное меню\n\nВыберите действие кнопкой в меню.", reply_markup=main_menu_keyboard())


async def main() -> None:
    load_dotenv()
    token = os.getenv("BOT_TOKEN")
    if not token:
        raise RuntimeError("BOT_TOKEN is missing. Put it into .env")

    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s %(levelname)s %(name)s: %(message)s",
    )
    logging.getLogger("aiogram.event").setLevel(logging.WARNING)
    ensure_storage()

    bot = Bot(token=token)
    dispatcher = Dispatcher(storage=MemoryStorage())
    dispatcher.include_router(router)
    await dispatcher.start_polling(bot)


if __name__ == "__main__":
    asyncio.run(main())
