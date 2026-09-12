import sqlite3
from pathlib import Path
from typing import Optional


BASE_DIR = Path(__file__).resolve().parent
DB_PATH = BASE_DIR / "wijze_lessons.db"


def get_connection():
    connection = sqlite3.connect(
        DB_PATH,
        check_same_thread=False
    )

    connection.row_factory = sqlite3.Row

    return connection


def init_database():
    connection = get_connection()

    cursor = connection.cursor()

    cursor.execute(
        """
        CREATE TABLE IF NOT EXISTS lessons (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            title TEXT NOT NULL,
            class_name TEXT DEFAULT '',
            grade TEXT DEFAULT '',
            study_direction TEXT DEFAULT '',
            duration INTEGER DEFAULT 50,
            topic TEXT DEFAULT '',
            learning_objectives TEXT DEFAULT '',
            notes TEXT DEFAULT '',
            phases_json TEXT DEFAULT '[]',
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL
        )
        """
    )

    connection.commit()
    connection.close()


def create_lesson(
    title: str,
    class_name: str,
    grade: str,
    study_direction: str,
    duration: int,
    topic: str,
    learning_objectives: str,
    notes: str,
    phases_json: str,
    created_at: str,
):
    connection = get_connection()

    cursor = connection.cursor()

    cursor.execute(
        """
        INSERT INTO lessons (
            title,
            class_name,
            grade,
            study_direction,
            duration,
            topic,
            learning_objectives,
            notes,
            phases_json,
            created_at,
            updated_at
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (
            title,
            class_name,
            grade,
            study_direction,
            duration,
            topic,
            learning_objectives,
            notes,
            phases_json,
            created_at,
            created_at,
        )
    )

    lesson_id = cursor.lastrowid

    connection.commit()
    connection.close()

    return lesson_id


def update_lesson(
    lesson_id: int,
    title: str,
    class_name: str,
    grade: str,
    study_direction: str,
    duration: int,
    topic: str,
    learning_objectives: str,
    notes: str,
    phases_json: str,
    updated_at: str,
):
    connection = get_connection()

    cursor = connection.cursor()

    cursor.execute(
        """
        UPDATE lessons
        SET
            title = ?,
            class_name = ?,
            grade = ?,
            study_direction = ?,
            duration = ?,
            topic = ?,
            learning_objectives = ?,
            notes = ?,
            phases_json = ?,
            updated_at = ?
        WHERE id = ?
        """,
        (
            title,
            class_name,
            grade,
            study_direction,
            duration,
            topic,
            learning_objectives,
            notes,
            phases_json,
            updated_at,
            lesson_id,
        )
    )

    connection.commit()
    connection.close()


def get_lessons():
    connection = get_connection()

    cursor = connection.cursor()

    cursor.execute(
        """
        SELECT *
        FROM lessons
        ORDER BY updated_at DESC
        """
    )

    rows = cursor.fetchall()

    connection.close()

    return rows


def get_lesson(lesson_id: int) -> Optional[sqlite3.Row]:
    connection = get_connection()

    cursor = connection.cursor()

    cursor.execute(
        """
        SELECT *
        FROM lessons
        WHERE id = ?
        """,
        (lesson_id,)
    )

    row = cursor.fetchone()

    connection.close()

    return row


def delete_lesson(lesson_id: int):
    connection = get_connection()

    cursor = connection.cursor()

    cursor.execute(
        """
        DELETE FROM lessons
        WHERE id = ?
        """,
        (lesson_id,)
    )

    connection.commit()
    connection.close()