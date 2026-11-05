"""培训机构管理系统 —— 学生档案模块

功能：学生档案的建立（增）、查询（含关键字搜索）、修改、删除
数据库：SQLite（自动在脚本目录生成 students.db）
访问方式：局域网内任意设备浏览器打开 http://<服务器IP>:5000
"""
import io
import os
import re
import shutil
import socket
import sqlite3
import time
import uuid
from datetime import datetime, timedelta
from functools import wraps

from flask import Flask, abort, jsonify, redirect, render_template, request, send_file, session, url_for
from werkzeug.security import check_password_hash, generate_password_hash

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(BASE_DIR, "students.db")
# 数据备份专用目录（与代码备份 code_backup_* 分开存放）
BACKUP_DIR = os.path.join(BASE_DIR, "backups", "data")
UPLOAD_DIR = os.path.join(BASE_DIR, "uploads", "photos")   # 照片存储目录

# ---------- Flask 应用 ----------
app = Flask(__name__)
app.secret_key = os.environ.get("FLASK_SECRET_KEY", "training-center-mgmt-2026-secret-key")
app.config["SESSION_PERMANENT"] = True
app.config["PERMANENT_SESSION_LIFETIME"] = timedelta(hours=8)
app.config["MAX_CONTENT_LENGTH"] = 60 * 1024 * 1024  # 60MB
app.config["TEMPLATES_AUTO_RELOAD"] = True   # 模板文件改动后无需重启即生效


# 角色定义：key 为 role 取值，value 为界面显示名
ROLES = {
    "admin": "管理员",
    "principal": "校长",
    "staff": "教务",
    "teacher": "教师",
    "viewer": "只读",
}

# 全系统唯一角色：每个角色有且只能存在一个账号（管理员 1 个、校长 1 个）
UNIQUE_ROLES = ("admin", "principal")

# 超级权限角色：可查看并管理系统的全部功能
SUPER_ROLES = ("admin", "principal")

# 学生档案「新增」权限：管理员 / 校长 / 教务 / 教师（只读角色不可新增）
STUDENT_CREATE_ROLES = ("admin", "principal", "staff", "teacher")

# 学生档案「修改 / 删除」权限：管理员 / 校长 / 教务（教师与只读不可）
STUDENT_WRITE_ROLES = ("admin", "principal", "staff")

# 学生简历「编辑」权限：管理员 / 校长 / 教务 / 教师（只读不可）
RESUME_WRITE_ROLES = ("admin", "principal", "staff", "teacher")

# ---------- 各业务模块权限点 ----------
# 科目管理：教务可新增 / 修改；删除仅管理员与校长（教务不可删除所创建的科目）
SUBJECT_WRITE_ROLES = ("admin", "principal", "staff")
SUBJECT_DELETE_ROLES = ("admin", "principal")

# 缴费管理：教务可新增 / 修改 / 导出；删除仅管理员与校长；教师无写入与导出
PAYMENT_WRITE_ROLES = ("admin", "principal", "staff")
PAYMENT_DELETE_ROLES = ("admin", "principal")
PAYMENT_EXPORT_ROLES = ("admin", "principal", "staff")

# 教师档案：教务仅可新建；修改 / 删除仅管理员与校长
TEACHER_CREATE_ROLES = ("admin", "principal", "staff")
TEACHER_WRITE_ROLES = ("admin", "principal")

# 考勤管理：教师可保存与补考勤；删除与导出仅管理员 / 校长 / 教务
ATTENDANCE_SAVE_ROLES = ("admin", "principal", "staff", "teacher")
ATTENDANCE_DELETE_ROLES = ("admin", "principal", "staff")
ATTENDANCE_EXPORT_ROLES = ("admin", "principal", "staff")
# 批量修改历史考勤（记录查询页「修改查询信息状态」）：仅管理员 / 校长 / 教务，教师不参与
ATTENDANCE_BATCH_ROLES = ("admin", "principal", "staff")

# 课时管理：导出仅管理员 / 校长 / 教务
LESSON_EXPORT_ROLES = ("admin", "principal", "staff")

# 系统设置：操作日志管理员 / 校长 / 教务 / 教师可查（其余功能仍仅管理员 / 校长）
SETTINGS_LOGS_ROLES = ("admin", "principal", "staff", "teacher")

# 课程管理：新建 / 编辑 / 删除课程（教师可协助维护，只读角色仅可查看）
COURSE_WRITE_ROLES = ("admin", "principal", "staff", "teacher")

# 报表统计：仅管理员与校长可查看
REPORT_VIEW_ROLES = ("admin", "principal")

# 退费管理：教务可新增 / 修改 / 导出；删除仅管理员与校长；教师无写入与导出
REFUND_WRITE_ROLES = ("admin", "principal", "staff")
REFUND_DELETE_ROLES = ("admin", "principal")
REFUND_EXPORT_ROLES = ("admin", "principal", "staff")

# 数据导入：仅管理员 / 校长 / 教务
IMPORT_ROLES = ("admin", "principal", "staff")

# 权限点 -> 允许的角色（前端 can() 与后端装饰器共用同一份定义）
PERMISSION_MATRIX = {
    "subject.write": SUBJECT_WRITE_ROLES,
    "subject.delete": SUBJECT_DELETE_ROLES,
    "payment.write": PAYMENT_WRITE_ROLES,
    "payment.delete": PAYMENT_DELETE_ROLES,
    "payment.export": PAYMENT_EXPORT_ROLES,
    "teacher.create": TEACHER_CREATE_ROLES,
    "teacher.write": TEACHER_WRITE_ROLES,
    "attendance.save": ATTENDANCE_SAVE_ROLES,
    "attendance.delete": ATTENDANCE_DELETE_ROLES,
    "attendance.export": ATTENDANCE_EXPORT_ROLES,
    "attendance.batch": ATTENDANCE_BATCH_ROLES,
    "course.write": COURSE_WRITE_ROLES,
    "lesson.export": LESSON_EXPORT_ROLES,
    "report.view": REPORT_VIEW_ROLES,
    "refund.write": REFUND_WRITE_ROLES,
    "refund.delete": REFUND_DELETE_ROLES,
    "refund.export": REFUND_EXPORT_ROLES,
    "import.all": IMPORT_ROLES,
    "student.create": STUDENT_CREATE_ROLES,
    "student.write": STUDENT_WRITE_ROLES,
    "resume.write": RESUME_WRITE_ROLES,
    "settings.manage": SUPER_ROLES,
    "settings.logs": SETTINGS_LOGS_ROLES,
}


def can(action):
    """当前登录用户是否具备指定权限点（模板与后端共用）"""
    allowed = PERMISSION_MATRIX.get(action)
    if allowed is None:
        return False
    return session.get("role") in allowed


def role_name(role):
    """角色名称映射"""
    return ROLES.get(role, role or "")


def is_super(role):
    """是否拥有全部功能管理权限（管理员 / 校长）"""
    return role in SUPER_ROLES


def can_create_student(role):
    """是否可以新增学生档案（教师也可新增）"""
    return role in STUDENT_CREATE_ROLES


def can_write_student(role):
    """是否可以修改 / 删除学生档案（仅管理员、校长、教务）"""
    return role in STUDENT_WRITE_ROLES


def can_write_resume(role):
    """是否可以编辑学生简历（管理员、校长、教务、教师）"""
    return role in RESUME_WRITE_ROLES


app.jinja_env.globals["role_name"] = role_name
app.jinja_env.globals["is_super"] = is_super
app.jinja_env.globals["can_create_student"] = can_create_student
app.jinja_env.globals["can_write_student"] = can_write_student
app.jinja_env.globals["can_write_resume"] = can_write_resume
app.jinja_env.globals["ROLES"] = ROLES
app.jinja_env.globals["can"] = can

# 照片限制
MAX_PHOTOS = 5                                  # 每个学生最多照片数
MAX_PHOTO_SIZE = 5 * 1024 * 1024                # 单张最大 5MB
ALLOWED_PHOTO_EXT = {".jpg", ".jpeg", ".png", ".gif", ".webp"}

# 简历条目类型
RESUME_TYPES = {
    "education": "教育经历",
    "training": "培训经历",
    "project": "项目经验",
    # 原「证书获取 certificate」与「竞赛获奖 award」合并为「获奖与荣誉 award」
    "award": "获奖与荣誉",
    "social": "社会实践与活动经历",
    "work": "作品展示",
}

# 「兴趣爱好与性格特点」组内的纯文字项：每项只保留一条内容，仅文字描述、不支持资料上传
# 自我评价复用 students.notes，其余两项存于 resume_items（description 字段）
TEXT_RESUME_TYPES = {
    "hobby": "兴趣爱好",
    "personality": "性格描述",
}

# 「获奖与荣誉」由旧的 certificate / award 两种 type 合并而来，历史数据统一改写为 award
LEGACY_RESUME_TYPE_MERGE = {"certificate": "award", "award": "award"}

# 简历栏目分组（决定编辑页与简历页的栏目顺序）
RESUME_GROUPS = [
    ("教育经历", ["education"]),
    ("培训经历", ["training"]),
    ("项目经验", ["project"]),
    ("获奖与荣誉", ["award"]),
    ("社会实践与活动经历", ["social"]),
    ("兴趣爱好与性格特点", ["hobby", "personality", "self_evaluation"]),
    ("作品展示", ["work"]),
]

# 简历条目资料附件（支持任意格式文件上传与下载）
ATTACH_DIR = os.path.join(BASE_DIR, "uploads", "files")     # 附件存储目录
MAX_ATT_SIZE = 50 * 1024 * 1024                 # 单个附件最大 50MB
MAX_ATTACHMENTS = 10                            # 每个条目最多附件数

# 学生档案字段（可编辑部分，监护人信息单独存于 guardians 表）
FIELDS = [
    "name", "gender", "birthdate", "id_card", "phone",
    "school", "grade", "course", "address", "notes",
]

# 18 位身份证号格式（末位可为 X）
ID_CARD_RE = re.compile(r"^\d{17}[\dXx]$")


# ---------- 数据库 ----------
def get_db():
    """获取数据库连接（Row 格式，便于转字典），启用 WAL 模式与外键"""
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA foreign_keys=ON")
    return conn


def next_student_no(conn):
    """生成下一个学生编号：字母 a-z 顺序 + 数字 00001-99999 顺序，如 a00001、a99999、b00001"""
    conn.execute("INSERT OR IGNORE INTO counters (name, value) VALUES ('student_seq', 0)")
    seq = conn.execute(
        "SELECT value FROM counters WHERE name = 'student_seq'"
    ).fetchone()["value"] + 1
    if seq > 26 * 99999:
        raise ValueError("学生编号已用尽")
    conn.execute(
        "UPDATE counters SET value = ? WHERE name = 'student_seq'", (seq,)
    )
    letter = chr(97 + (seq - 1) // 99999)          # a-z
    number = (seq - 1) % 99999 + 1                  # 1-99999
    return f"{letter}{number:05d}"


def next_teacher_no(conn):
    """生成下一个教师工号：字母 a-z 顺序 + 数字 0001-9999 顺序，如 a0001、a9999、b0001"""
    conn.execute("INSERT OR IGNORE INTO counters (name, value) VALUES ('teacher_seq', 0)")
    seq = conn.execute(
        "SELECT value FROM counters WHERE name = 'teacher_seq'"
    ).fetchone()["value"] + 1
    if seq > 26 * 9999:
        raise ValueError("教师工号已用尽")
    conn.execute(
        "UPDATE counters SET value = ? WHERE name = 'teacher_seq'", (seq,)
    )
    letter = chr(97 + (seq - 1) // 9999)           # a-z
    number = (seq - 1) % 9999 + 1                  # 1-9999
    return f"{letter}{number:04d}"


def init_db():
    """初始化数据表（不存在时创建），并兼容旧版本数据结构"""
    os.makedirs(UPLOAD_DIR, exist_ok=True)
    os.makedirs(ATTACH_DIR, exist_ok=True)
    conn = get_db()
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS students (
            id              INTEGER PRIMARY KEY AUTOINCREMENT,
            student_no      TEXT UNIQUE,
            name            TEXT NOT NULL,
            gender          TEXT NOT NULL DEFAULT '男',
            birthdate       TEXT,
            id_card         TEXT,
            phone           TEXT,
            school          TEXT,
            grade           TEXT,
            course          TEXT,
            address         TEXT,
            notes           TEXT,
            created_at      TEXT NOT NULL,
            updated_at      TEXT
        )
        """
    )
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS guardians (
            id          INTEGER PRIMARY KEY AUTOINCREMENT,
            student_id  INTEGER NOT NULL,
            name        TEXT,
            phone       TEXT,
            wechat      TEXT
        )
        """
    )
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS photos (
            id              INTEGER PRIMARY KEY AUTOINCREMENT,
            student_id      INTEGER NOT NULL,
            filename        TEXT NOT NULL,
            original_name   TEXT,
            uploaded_at     TEXT,
            sort            INTEGER NOT NULL DEFAULT 0
        )
        """
    )
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS resume_items (
            id          INTEGER PRIMARY KEY AUTOINCREMENT,
            student_id  INTEGER NOT NULL,
            type        TEXT NOT NULL,
            title       TEXT NOT NULL,
            subtitle    TEXT,
            date_text   TEXT,
            description TEXT,
            created_at  TEXT
        )
        """
    )
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS attachments (
            id              INTEGER PRIMARY KEY AUTOINCREMENT,
            student_id      INTEGER NOT NULL,
            item_id         INTEGER NOT NULL,
            filename        TEXT NOT NULL,
            original_name   TEXT,
            size            INTEGER,
            uploaded_at     TEXT
        )
        """
    )
    # 历史数据迁移：旧的 certificate（证书获取）条目并入 award（获奖与荣誉）
    for old_type, new_type in LEGACY_RESUME_TYPE_MERGE.items():
        if old_type != new_type:
            conn.execute(
                "UPDATE resume_items SET type = ? WHERE type = ?", (new_type, old_type)
            )
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS attendance (
            id          INTEGER PRIMARY KEY AUTOINCREMENT,
            student_id  INTEGER NOT NULL,
            date        TEXT NOT NULL,
            status      TEXT NOT NULL,
            remark      TEXT,
            course_name TEXT,
            lesson      TEXT,
            teacher     TEXT,
            created_at  TEXT,
            updated_at  TEXT,
            UNIQUE (student_id, date)
        )
        """
    )
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS payment_photos (
            id              INTEGER PRIMARY KEY AUTOINCREMENT,
            payment_id      INTEGER NOT NULL,
            filename        TEXT NOT NULL,
            original_name   TEXT,
            size            INTEGER,
            uploaded_at     TEXT,
            data            BLOB
        )
        """
    )
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS refund_photos (
            id              INTEGER PRIMARY KEY AUTOINCREMENT,
            refund_id       INTEGER NOT NULL,
            filename        TEXT NOT NULL,
            original_name   TEXT,
            size            INTEGER,
            uploaded_at     TEXT,
            data            BLOB
        )
        """
    )
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS courses (
            id          INTEGER PRIMARY KEY AUTOINCREMENT,
            name        TEXT NOT NULL,
            category    TEXT,
            teacher     TEXT,
            schedule    TEXT,
            hours       INTEGER,
            price       REAL,
            capacity    INTEGER,
            start_date  TEXT,
            end_date    TEXT,
            description TEXT,
            status      TEXT NOT NULL DEFAULT '开课中',
            created_at  TEXT
        )
        """
    )
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS enrollments (
            id          INTEGER PRIMARY KEY AUTOINCREMENT,
            course_id   INTEGER NOT NULL,
            student_id  INTEGER NOT NULL,
            created_at  TEXT,
            UNIQUE (course_id, student_id)
        )
        """
    )
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS teachers (
            id          INTEGER PRIMARY KEY AUTOINCREMENT,
            teacher_no  TEXT UNIQUE,
            name        TEXT NOT NULL,
            gender      TEXT NOT NULL DEFAULT '男',
            phone       TEXT,
            subject     TEXT,
            title       TEXT,
            hire_date   TEXT,
            status      TEXT NOT NULL DEFAULT '在职',
            notes       TEXT,
            created_at  TEXT
        )
        """
    )
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS payments (
            id          INTEGER PRIMARY KEY AUTOINCREMENT,
            student_id  INTEGER NOT NULL,
            course      TEXT,
            amount      REAL NOT NULL,
            pay_date    TEXT,
            method      TEXT,
            remark      TEXT,
            created_at  TEXT
        )
        """
    )
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS subjects (
            id           INTEGER PRIMARY KEY AUTOINCREMENT,
            name         TEXT NOT NULL,
            price        REAL NOT NULL DEFAULT 0,
            lesson_count INTEGER NOT NULL DEFAULT 1,
            lesson_price REAL NOT NULL DEFAULT 0,
            created_at   TEXT,
            updated_at   TEXT
        )
        """
    )
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS counters (
            name  TEXT PRIMARY KEY,
            value INTEGER NOT NULL DEFAULT 0
        )
        """
    )

    # 旧库迁移：photos 表补充排序列（用于指定证件照）
    pcols = {row["name"] for row in conn.execute("PRAGMA table_info(photos)")}
    if "sort" not in pcols:
        conn.execute("ALTER TABLE photos ADD COLUMN sort INTEGER NOT NULL DEFAULT 0")

    # 旧库迁移：payments 表补充计价字段
    paycols = {row["name"] for row in conn.execute("PRAGMA table_info(payments)")}
    if "unit_price" not in paycols:
        conn.execute("ALTER TABLE payments ADD COLUMN unit_price REAL DEFAULT 0")
    if "lesson_count" not in paycols:
        conn.execute("ALTER TABLE payments ADD COLUMN lesson_count INTEGER DEFAULT 0")
    if "discount_percent" not in paycols:
        conn.execute("ALTER TABLE payments ADD COLUMN discount_percent REAL DEFAULT 100")
    if "discount_amount" not in paycols:
        conn.execute("ALTER TABLE payments ADD COLUMN discount_amount REAL DEFAULT 0")

    # 旧库迁移：补充新增字段
    cols = {row["name"] for row in conn.execute("PRAGMA table_info(students)")}
    if "id_card" not in cols:
        conn.execute("ALTER TABLE students ADD COLUMN id_card TEXT")
    if "updated_at" not in cols:
        conn.execute("ALTER TABLE students ADD COLUMN updated_at TEXT")
    if "student_no" not in cols:
        conn.execute("ALTER TABLE students ADD COLUMN student_no TEXT")
    if "grade" not in cols:
        conn.execute("ALTER TABLE students ADD COLUMN grade TEXT")

    # 旧库迁移：为已有学生按建档顺序回填六位编号（a00001 起）
    for row in conn.execute(
        "SELECT id FROM students WHERE student_no IS NULL OR student_no = '' ORDER BY id"
    ).fetchall():
        no = next_student_no(conn)
        conn.execute(
            "UPDATE students SET student_no = ? WHERE id = ?", (no, row["id"])
        )
    conn.execute(
        "CREATE UNIQUE INDEX IF NOT EXISTS idx_students_student_no ON students(student_no)"
    )

    # 旧库迁移：补充教师工号列，并为已有教师按建档顺序回填（a0001 起）
    if "teacher_no" not in {
        row["name"] for row in conn.execute("PRAGMA table_info(teachers)")
    }:
        conn.execute("ALTER TABLE teachers ADD COLUMN teacher_no TEXT")
    for row in conn.execute(
        "SELECT id FROM teachers WHERE teacher_no IS NULL OR teacher_no = '' ORDER BY id"
    ).fetchall():
        no = next_teacher_no(conn)
        conn.execute(
            "UPDATE teachers SET teacher_no = ? WHERE id = ?", (no, row["id"])
        )
    conn.execute(
        "CREATE UNIQUE INDEX IF NOT EXISTS idx_teachers_teacher_no ON teachers(teacher_no)"
    )

    # 旧库迁移：补充课程开课/结课日期列
    ccols = {row["name"] for row in conn.execute("PRAGMA table_info(courses)")}
    if "start_date" not in ccols:
        conn.execute("ALTER TABLE courses ADD COLUMN start_date TEXT")
    if "end_date" not in ccols:
        conn.execute("ALTER TABLE courses ADD COLUMN end_date TEXT")
    # 已有课程：开课日期默认取创建当天（与「开课日期默认为创建时间」规则一致），结课日期留空
    conn.execute(
        """
        UPDATE courses SET start_date = substr(created_at, 1, 10)
        WHERE (start_date IS NULL OR start_date = '')
          AND created_at IS NOT NULL AND created_at != ''
        """
    )

    # 旧库迁移：考勤表补充课程名称 / 本节课名称 / 老师姓名列
    acols = {row["name"] for row in conn.execute("PRAGMA table_info(attendance)")}
    if "course_name" not in acols:
        conn.execute("ALTER TABLE attendance ADD COLUMN course_name TEXT")
    if "lesson" not in acols:
        conn.execute("ALTER TABLE attendance ADD COLUMN lesson TEXT")
    if "teacher" not in acols:
        conn.execute("ALTER TABLE attendance ADD COLUMN teacher TEXT")
    # v3.0 迁移：考勤表补充上课时间段与来源列（补考勤 / 实际课时）
    if "start_time" not in acols:
        conn.execute("ALTER TABLE attendance ADD COLUMN start_time TEXT")
    if "end_time" not in acols:
        conn.execute("ALTER TABLE attendance ADD COLUMN end_time TEXT")
    if "source" not in acols:
        conn.execute("ALTER TABLE attendance ADD COLUMN source TEXT DEFAULT ''")

    # 旧库迁移：把旧的监护人两个字段搬进 guardians 表
    if "guardian_name" in cols or "guardian_phone" in cols:
        legacy = conn.execute(
            """
            SELECT id, guardian_name, guardian_phone FROM students
            WHERE COALESCE(guardian_name, '') != '' OR COALESCE(guardian_phone, '') != ''
            """
        ).fetchall()
        for row in legacy:
            exists = conn.execute(
                "SELECT 1 FROM guardians WHERE student_id = ?", (row["id"],)
            ).fetchone()
            if not exists:
                conn.execute(
                    "INSERT INTO guardians (student_id, name, phone, wechat) VALUES (?, ?, ?, '')",
                    (row["id"], row["guardian_name"] or "", row["guardian_phone"] or ""),
                )

    # 迁移：photos 表增加 data BLOB 列，将磁盘照片读入数据库
    pcols2 = {row["name"] for row in conn.execute("PRAGMA table_info(photos)")}
    if "data" not in pcols2:
        conn.execute("ALTER TABLE photos ADD COLUMN data BLOB")
    for row in conn.execute(
        "SELECT id, filename FROM photos WHERE data IS NULL AND filename IS NOT NULL"
    ).fetchall():
        path = os.path.join(UPLOAD_DIR, row["filename"])
        if os.path.exists(path):
            with open(path, "rb") as f:
                conn.execute("UPDATE photos SET data = ? WHERE id = ?", (f.read(), row["id"]))

    # 迁移：attachments 表增加 data BLOB 列，将磁盘附件读入数据库
    att_cols = {row["name"] for row in conn.execute("PRAGMA table_info(attachments)")}
    if "data" not in att_cols:
        conn.execute("ALTER TABLE attachments ADD COLUMN data BLOB")
    for row in conn.execute(
        "SELECT id, filename FROM attachments WHERE data IS NULL AND filename IS NOT NULL"
    ).fetchall():
        path = os.path.join(ATTACH_DIR, row["filename"])
        if os.path.exists(path):
            with open(path, "rb") as f:
                conn.execute("UPDATE attachments SET data = ? WHERE id = ?", (f.read(), row["id"]))

    # ---- v2.0 新增表 ----
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS users (
            id              INTEGER PRIMARY KEY AUTOINCREMENT,
            username        TEXT UNIQUE NOT NULL,
            password_hash   TEXT NOT NULL,
            role            TEXT NOT NULL DEFAULT 'staff',
            display_name    TEXT,
            status          TEXT NOT NULL DEFAULT '启用',
            failed_attempts INTEGER NOT NULL DEFAULT 0,
            locked_until    TEXT,
            last_login      TEXT,
            created_at      TEXT
        )
        """
    )
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS operation_logs (
            id          INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id     INTEGER,
            username    TEXT,
            action      TEXT,
            module      TEXT,
            target      TEXT,
            detail      TEXT,
            ip          TEXT,
            created_at  TEXT
        )
        """
    )
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS refunds (
            id          INTEGER PRIMARY KEY AUTOINCREMENT,
            student_id  INTEGER NOT NULL,
            payment_id  INTEGER,
            amount      REAL NOT NULL,
            lesson_count INTEGER DEFAULT 0,
            refund_date TEXT,
            method      TEXT DEFAULT '原路退回',
            reason      TEXT,
            remark      TEXT,
            created_at  TEXT
        )
        """
    )
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS system_settings (
            key         TEXT PRIMARY KEY,
            value       TEXT,
            updated_at  TEXT
        )
        """

    )
    # v2.0 迁移：courses 表补充 teacher_id 外键列
    ccols2 = {row["name"] for row in conn.execute("PRAGMA table_info(courses)")}
    if "teacher_id" not in ccols2:
        conn.execute("ALTER TABLE courses ADD COLUMN teacher_id INTEGER")
    if "description" not in ccols2:
        conn.execute("ALTER TABLE courses ADD COLUMN description TEXT")

    # v2.0 迁移：payments 表补充 status 列
    paycols2 = {row["name"] for row in conn.execute("PRAGMA table_info(payments)")}
    if "status" not in paycols2:
        conn.execute("ALTER TABLE payments ADD COLUMN status TEXT DEFAULT '已确认'")

    # v2.0 迁移：refunds 表补充 lesson_count（退课时数）列
    rcols2 = {row["name"] for row in conn.execute("PRAGMA table_info(refunds)")}
    if "lesson_count" not in rcols2:
        conn.execute("ALTER TABLE refunds ADD COLUMN lesson_count INTEGER DEFAULT 0")

    # v2.0 迁移：teachers 表补充 notes 列
    tcols = {row["name"] for row in conn.execute("PRAGMA table_info(teachers)")}
    if "notes" not in tcols:
        conn.execute("ALTER TABLE teachers ADD COLUMN notes TEXT")

    # v2.0 索引优化
    conn.execute("CREATE INDEX IF NOT EXISTS idx_attendance_student ON attendance(student_id)")
    conn.execute("CREATE INDEX IF NOT EXISTS idx_payments_student ON payments(student_id)")
    conn.execute("CREATE INDEX IF NOT EXISTS idx_enrollments_course ON enrollments(course_id)")
    conn.execute("CREATE INDEX IF NOT EXISTS idx_enrollments_student ON enrollments(student_id)")
    conn.execute("CREATE INDEX IF NOT EXISTS idx_operation_logs_created ON operation_logs(created_at)")
    conn.execute("CREATE INDEX IF NOT EXISTS idx_payment_photos_payment ON payment_photos(payment_id)")
    conn.execute("CREATE INDEX IF NOT EXISTS idx_refund_photos_refund ON refund_photos(refund_id)")

    # v2.0 初始化默认管理员账号
    admin_exists = conn.execute("SELECT 1 FROM users WHERE username = 'admin'").fetchone()
    if not admin_exists:
        conn.execute(
            "INSERT INTO users (username, password_hash, role, display_name, status, created_at) "
            "VALUES (?, ?, 'admin', '系统管理员', '启用', ?)",
            ("admin", generate_password_hash("admin123"), datetime.now().strftime("%Y-%m-%d %H:%M:%S")),
        )

    # 初始化默认校长账号（管理员与校长有且只能各一个）
    principal_exists = conn.execute(
        "SELECT 1 FROM users WHERE role = 'principal' LIMIT 1"
    ).fetchone()
    if not principal_exists:
        conn.execute(
            "INSERT INTO users (username, password_hash, role, display_name, status, created_at) "
            "VALUES (?, ?, 'principal', '校长', '启用', ?)",
            ("principal", generate_password_hash("principal123"),
             datetime.now().strftime("%Y-%m-%d %H:%M:%S")),
        )

    # v2.0 初始化默认系统设置
    default_settings = {
        "org_name": "培训机构",
        "attendance_warning_threshold": "3",
        "session_timeout_minutes": "480",
    }
    for k, v in default_settings.items():
        exists = conn.execute("SELECT 1 FROM system_settings WHERE key = ?", (k,)).fetchone()
        if not exists:
            conn.execute(
                "INSERT INTO system_settings (key, value, updated_at) VALUES (?, ?, ?)",
                (k, v, datetime.now().strftime("%Y-%m-%d %H:%M:%S")),
            )

    conn.commit()
    conn.close()


def clean_payload(data):
    """提取并清理前端提交的数据"""
    payload = {f: str(data.get(f) or "").strip() for f in FIELDS}
    if payload["gender"] not in ("男", "女"):
        payload["gender"] = "男"
    return payload


def extract_guardians(data):
    """提取并清理监护人列表（姓名/电话/微信均为空的组不保留）"""
    result = []
    for g in data.get("guardians") or []:
        if not isinstance(g, dict):
            continue
        item = {
            "name": str(g.get("name") or "").strip(),
            "phone": str(g.get("phone") or "").strip(),
            "wechat": str(g.get("wechat") or "").strip(),
        }
        if any(item.values()):
            result.append(item)
    return result


def validate(payload):
    """基础校验，返回错误信息（None 表示通过）"""
    if not payload["name"]:
        return "姓名为必填项"
    if payload["id_card"] and not ID_CARD_RE.match(payload["id_card"]):
        return "身份证号格式不正确（应为18位）"
    return None


# ---------- 工具 ----------
def get_lan_ip():
    """获取本机局域网 IP"""
    s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    try:
        s.connect(("8.8.8.8", 80))  # 不会真正发送数据，仅用于确定路由
        return s.getsockname()[0]
    except OSError:
        return "127.0.0.1"
    finally:
        s.close()


def safe_remove(path):
    """删除磁盘文件（Windows 下文件句柄可能延迟释放，自动重试）"""
    for _ in range(10):
        try:
            os.remove(path)
            return True
        except FileNotFoundError:
            return False
        except OSError:
            time.sleep(0.1)
    return False


# ---------- v2.0 认证与权限 ----------
def current_user(conn=None):
    """获取当前登录用户，返回 dict 或 None"""
    uid = session.get("user_id")
    if not uid:
        return None
    own_conn = conn is None
    if own_conn:
        conn = get_db()
    row = conn.execute("SELECT * FROM users WHERE id = ?", (uid,)).fetchone()
    if own_conn:
        conn.close()
    if not row:
        return None
    return dict(row)


def _viewer_download_denied():
    """只读角色禁止以附件方式下载文件（在线预览不受限）"""
    return session.get("role") == "viewer" and bool(request.args.get("download"))


def log_operation(conn, action, module, target="", detail=""):
    """记录操作审计日志"""
    u = current_user(conn)
    username = u["username"] if u else "anonymous"
    user_id = u["id"] if u else None
    ip = request.remote_addr or ""
    conn.execute(
        "INSERT INTO operation_logs (user_id, username, action, module, target, detail, ip, created_at) "
        "VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
        (user_id, username, action, module, target, detail, ip,
         datetime.now().strftime("%Y-%m-%d %H:%M:%S")),
    )


def login_required(f):
    """装饰器：检查是否登录"""
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if not session.get("user_id"):
            if request.path.startswith("/api/"):
                return jsonify({"error": "未登录", "redirect": "/login"}), 401
            return redirect("/login")
        return f(*args, **kwargs)
    return decorated_function


def super_required(f):
    """装饰器：检查是否超级权限角色（管理员 / 校长），二者可管理全部功能"""
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if not session.get("user_id"):
            if request.path.startswith("/api/"):
                return jsonify({"error": "未登录", "redirect": "/login"}), 401
            return redirect("/login")
        if not is_super(session.get("role")):
            if request.path.startswith("/api/"):
                return jsonify({"error": "权限不足"}), 403
            return redirect("/")
        return f(*args, **kwargs)
    return decorated_function


# 兼容旧名称：现语义为「管理员 / 校长」均可访问
admin_required = super_required


def _role_required(allowed, message="权限不足"):
    """角色白名单装饰器工厂：登录校验 + 角色校验"""
    def decorator(f):
        @wraps(f)
        def decorated_function(*args, **kwargs):
            if not session.get("user_id"):
                if request.path.startswith("/api/"):
                    return jsonify({"error": "未登录", "redirect": "/login"}), 401
                return redirect("/login")
            if session.get("role") not in allowed:
                if request.path.startswith("/api/"):
                    return jsonify({"error": message}), 403
                return redirect("/")
            return f(*args, **kwargs)
        return decorated_function
    return decorator


def student_create_required(f):
    """学生档案新增权限：管理员 / 校长 / 教务 / 教师"""
    return _role_required(STUDENT_CREATE_ROLES, "当前角色无权新增学生档案")(f)


def student_write_required(f):
    """学生档案修改 / 删除权限：管理员 / 校长 / 教务"""
    return _role_required(STUDENT_WRITE_ROLES, "当前角色无权修改或删除学生档案")(f)


def resume_write_required(f):
    """学生简历编辑权限：管理员 / 校长 / 教务 / 教师"""
    return _role_required(RESUME_WRITE_ROLES, "当前角色无权编辑简历")(f)


def permission_required(action, message="权限不足"):
    """按权限点校验的装饰器（参数化形式）

    用法：@permission_required("payment.delete", "教务不可删除缴费记录")
    """
    return _role_required(PERMISSION_MATRIX.get(action, ()), message)


def get_setting(conn, key, default=""):
    """读取系统设置项"""
    row = conn.execute("SELECT value FROM system_settings WHERE key = ?", (key,)).fetchone()
    return row["value"] if row else default


# ---------- v2.0 认证中间件 ----------
OPEN_PATHS = {"/login", "/api/auth/login", "/static", "/uploads"}

@app.before_request
def check_auth():
    """全局请求拦截：检查登录状态（排除登录页和静态资源）"""
    path = request.path
    for op in OPEN_PATHS:
        if path == op or path.startswith(op + "/"):
            return None
    if not session.get("user_id"):
        if path.startswith("/api/"):
            return jsonify({"error": "未登录", "redirect": "/login"}), 401
        return redirect("/login")
    return None


# ---------- v2.0 认证路由 ----------
@app.route("/login")
def login_page():
    """登录页"""
    if session.get("user_id"):
        return redirect("/")
    return render_template("login.html")


@app.route("/logout")
def logout():
    """退出登录"""
    conn = get_db()
    try:
        log_operation(conn, "登出", "认证", session.get("username", ""))
        conn.commit()
    finally:
        conn.close()
    session.clear()
    return redirect("/login")


# ---------- v2.0 认证 API ----------
@app.post("/api/auth/login")
def api_login():
    """用户登录"""
    d = request.get_json(silent=True) or {}
    username = str(d.get("username") or "").strip()
    password = str(d.get("password") or "")
    if not username or not password:
        return jsonify({"error": "请输入用户名和密码"}), 400

    conn = get_db()
    try:
        row = conn.execute(
            "SELECT * FROM users WHERE username = ?", (username,)
        ).fetchone()
        if not row:
            return jsonify({"error": "用户名或密码错误"}), 401
        user = dict(row)

        # 检查账号锁定
        if user["locked_until"]:
            try:
                locked_until = datetime.strptime(user["locked_until"], "%Y-%m-%d %H:%M:%S")
                if locked_until > datetime.now():
                    remaining = int((locked_until - datetime.now()).total_seconds() / 60) + 1
                    return jsonify({"error": f"账号已锁定，请 {remaining} 分钟后重试"}), 403
            except ValueError:
                pass

        # 检查账号状态
        if user["status"] != "启用":
            return jsonify({"error": "账号已被禁用，请联系管理员"}), 403

        # 验证密码
        if not check_password_hash(user["password_hash"], password):
            attempts = user["failed_attempts"] + 1
            if attempts >= 5:
                locked_until = (datetime.now() + timedelta(minutes=15)).strftime("%Y-%m-%d %H:%M:%S")
                conn.execute(
                    "UPDATE users SET failed_attempts = 0, locked_until = ? WHERE id = ?",
                    (locked_until, user["id"]),
                )
                conn.commit()
                return jsonify({"error": "密码错误次数过多，账号已锁定 15 分钟"}), 403
            else:
                conn.execute(
                    "UPDATE users SET failed_attempts = ? WHERE id = ?",
                    (attempts, user["id"]),
                )
                conn.commit()
                remaining = 5 - attempts
                return jsonify({"error": f"用户名或密码错误（还剩 {remaining} 次尝试机会）"}), 401

        # 登录成功
        now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        conn.execute(
            "UPDATE users SET failed_attempts = 0, locked_until = NULL, last_login = ? WHERE id = ?",
            (now_str, user["id"]),
        )
        conn.commit()
    finally:
        conn.close()

    session.permanent = True
    session["user_id"] = user["id"]
    session["username"] = user["username"]
    session["role"] = user["role"]
    session["display_name"] = user["display_name"] or user["username"]

    # 登录日志需在 session 建立后写入，日志才能正确关联到该用户（user_id）
    conn = get_db()
    try:
        log_operation(conn, "登录", "认证", username)
        conn.commit()
    finally:
        conn.close()

    # 首次登录（last_login 为空）需要强制改密码
    force_change = user["last_login"] is None
    return jsonify({
        "message": "登录成功",
        "user": {
            "id": user["id"],
            "username": user["username"],
            "role": user["role"],
            "display_name": user["display_name"] or user["username"],
        },
        "force_change_password": force_change,
    })


@app.post("/api/auth/logout")
def api_logout():
    """API退出登录"""
    conn = get_db()
    try:
        log_operation(conn, "登出", "认证", session.get("username", ""))
        conn.commit()
    finally:
        conn.close()
    session.clear()
    return jsonify({"message": "已退出登录"})


@app.get("/api/auth/me")
def api_me():
    """获取当前登录用户信息"""
    u = current_user()
    if not u:
        return jsonify({"error": "未登录"}), 401
    return jsonify({
        "id": u["id"],
        "username": u["username"],
        "role": u["role"],
        "display_name": u["display_name"] or u["username"],
        "last_login": u["last_login"],
    })


@app.post("/api/auth/change-password")
def api_change_password():
    """修改密码"""
    d = request.get_json(silent=True) or {}
    old_password = str(d.get("old_password") or "")
    new_password = str(d.get("new_password") or "")
    if not old_password or not new_password:
        return jsonify({"error": "请输入原密码和新密码"}), 400
    if len(new_password) < 6:
        return jsonify({"error": "新密码至少 6 个字符"}), 400

    conn = get_db()
    try:
        u = current_user(conn)
        if not u:
            return jsonify({"error": "未登录"}), 401
        if not check_password_hash(u["password_hash"], old_password):
            return jsonify({"error": "原密码错误"}), 400
        conn.execute(
            "UPDATE users SET password_hash = ? WHERE id = ?",
            (generate_password_hash(new_password), u["id"]),
        )
        log_operation(conn, "修改密码", "认证", u["username"])
        conn.commit()
    finally:
        conn.close()
    return jsonify({"message": "密码修改成功"})


# ---------- v2.0 用户管理（管理员 / 校长） ----------
@app.route("/settings")
@login_required
def settings_page():
    """系统设置页：全部功能仅管理员 / 校长可见，其余角色仅可修改自己的密码"""
    conn = get_db()
    try:
        user = current_user()
        super_user = is_super((user or {}).get("role"))
        users = [dict(r) for r in conn.execute(
            "SELECT id, username, role, display_name, status, last_login, created_at FROM users ORDER BY id"
        ).fetchall()] if super_user else []
        settings = {r["key"]: r["value"] for r in conn.execute(
            "SELECT key, value FROM system_settings"
        ).fetchall()} if super_user else {}
        backups = list_backups() if super_user else []
    finally:
        conn.close()
    return render_template(
        "settings.html", users=users, settings=settings, backups=backups,
        user=user, super_user=super_user,
    )


@app.get("/api/users")
@super_required
def list_users():
    """用户列表"""
    conn = get_db()
    try:
        rows = conn.execute(
            "SELECT id, username, role, display_name, status, last_login, created_at FROM users ORDER BY id"
        ).fetchall()
    finally:
        conn.close()
    return jsonify([dict(r) for r in rows])


@app.post("/api/users")
@super_required
def create_user():
    """新建用户"""
    d = request.get_json(silent=True) or {}
    username = str(d.get("username") or "").strip()
    password = str(d.get("password") or "")
    role = str(d.get("role") or "").strip()
    display_name = str(d.get("display_name") or "").strip()
    if len(username) < 3:
        return jsonify({"error": "用户名至少 3 个字符"}), 400
    if len(password) < 6:
        return jsonify({"error": "密码至少 6 个字符"}), 400
    if role not in ROLES:
        role = "staff"
    conn = get_db()
    try:
        exists = conn.execute("SELECT 1 FROM users WHERE username = ?", (username,)).fetchone()
        if exists:
            return jsonify({"error": "用户名已存在"}), 400
        # 管理员 / 校长全系统各只能有一个
        if role in UNIQUE_ROLES:
            occupied = conn.execute(
                "SELECT username FROM users WHERE role = ? LIMIT 1", (role,)
            ).fetchone()
            if occupied:
                return jsonify({
                    "error": f"{ROLES[role]}有且只能有一个，已存在（{occupied['username']}）"
                }), 400
        now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        cur = conn.execute(
            "INSERT INTO users (username, password_hash, role, display_name, status, created_at) "
            "VALUES (?, ?, ?, ?, '启用', ?)",
            (username, generate_password_hash(password), role, display_name, now_str),
        )
        log_operation(conn, "创建用户", "用户管理", username)
        conn.commit()
        uid = cur.lastrowid
    finally:
        conn.close()
    return jsonify({"id": uid, "message": "用户创建成功"}), 201


@app.put("/api/users/<int:user_id>")
@super_required
def update_user(user_id):
    """修改用户（角色/显示名/状态/重置密码）"""
    d = request.get_json(silent=True) or {}
    conn = get_db()
    try:
        row = conn.execute("SELECT * FROM users WHERE id = ?", (user_id,)).fetchone()
        if not row:
            return jsonify({"error": "用户不存在"}), 404
        role = str(d.get("role") or row["role"] or "").strip()
        if role not in ROLES:
            role = row["role"]
        display_name = str(d.get("display_name") or row["display_name"] or "").strip()
        status = str(d.get("status") or row["status"] or "").strip()
        if status not in ("启用", "禁用"):
            status = row["status"]
        is_self = user_id == session.get("user_id")

        # 管理员 / 校长各只能有一个：目标角色已被他人占用时拒绝
        if role != row["role"] and role in UNIQUE_ROLES:
            occupied = conn.execute(
                "SELECT username FROM users WHERE role = ? AND id != ? LIMIT 1", (role, user_id)
            ).fetchone()
            if occupied:
                return jsonify({
                    "error": f"{ROLES[role]}有且只能有一个，已存在（{occupied['username']}）"
                }), 400

        # 不能通过改自己把唯一的管理员/校长降级或禁用，避免系统失去管理员
        if is_self and (role != row["role"] or status != "启用"):
            if row["role"] in UNIQUE_ROLES:
                return jsonify({"error": "不能修改自己的角色或状态"}), 400

        new_password = str(d.get("password") or "").strip()
        if new_password:
            if len(new_password) < 6:
                return jsonify({"error": "密码至少 6 个字符"}), 400
            conn.execute(
                "UPDATE users SET password_hash = ?, role = ?, display_name = ?, status = ? WHERE id = ?",
                (generate_password_hash(new_password), role, display_name, status, user_id),
            )
        else:
            conn.execute(
                "UPDATE users SET role = ?, display_name = ?, status = ? WHERE id = ?",
                (role, display_name, status, user_id),
            )
        log_operation(conn, "修改用户", "用户管理", row["username"])
        conn.commit()
    finally:
        conn.close()
    return jsonify({"message": "修改成功"})


@app.delete("/api/users/<int:user_id>")
@super_required
def delete_user(user_id):
    """删除用户"""
    conn = get_db()
    try:
        row = conn.execute("SELECT username, role FROM users WHERE id = ?", (user_id,)).fetchone()
        if not row:
            return jsonify({"error": "用户不存在"}), 404
        if row["username"] == "admin":
            return jsonify({"error": "不能删除默认管理员账号"}), 400
        # 管理员 / 校长唯一：持有者不可删除
        if row["role"] in UNIQUE_ROLES:
            return jsonify({"error": f"{ROLES[row['role']]}有且只能有一个，不可删除"}), 400
        if user_id == session.get("user_id"):
            return jsonify({"error": "不能删除当前登录账号"}), 400
        conn.execute("DELETE FROM users WHERE id = ?", (user_id,))
        # 操作日志跟随账户：删除用户的同时清掉其全部日志（含旧数据中仅按用户名记录的）
        conn.execute(
            "DELETE FROM operation_logs WHERE user_id = ? OR username = ?",
            (user_id, row["username"]),
        )
        log_operation(conn, "删除用户", "用户管理", row["username"])
        conn.commit()
    finally:
        conn.close()
    return jsonify({"message": "删除成功"})


# ---------- v2.0 备份功能 ----------
def _migrate_legacy_backups():
    """一次性迁移：把旧版散落在 backups/ 根目录的 .db 备份移到 backups/data/ 专用目录"""
    legacy_dir = os.path.dirname(BACKUP_DIR)   # backups/
    try:
        for f in os.listdir(legacy_dir):
            if f.endswith(".db"):
                src = os.path.join(legacy_dir, f)
                dst = os.path.join(BACKUP_DIR, f)
                if not os.path.exists(dst):
                    shutil.move(src, dst)
    except FileNotFoundError:
        pass


def list_backups():
    """列出所有备份文件（backups/data/ 专用目录）"""
    os.makedirs(BACKUP_DIR, exist_ok=True)
    _migrate_legacy_backups()
    result = []
    for f in sorted(os.listdir(BACKUP_DIR), reverse=True):
        if f.endswith(".db"):
            path = os.path.join(BACKUP_DIR, f)
            stat = os.stat(path)
            result.append({
                "filename": f,
                "size": stat.st_size,
                "created_at": datetime.fromtimestamp(stat.st_mtime).strftime("%Y-%m-%d %H:%M:%S"),
            })
    return result


def backup_database():
    """备份数据库到 backups/data/ 专用目录"""
    os.makedirs(BACKUP_DIR, exist_ok=True)
    filename = f"students_{datetime.now().strftime('%Y%m%d_%H%M%S')}.db"
    dest = os.path.join(BACKUP_DIR, filename)
    shutil.copy2(DB_PATH, dest)
    return filename


# ---------- v2.0 系统设置 API ----------
@app.get("/api/settings")
@super_required
def api_get_settings():
    """获取系统设置"""
    conn = get_db()
    try:
        rows = conn.execute("SELECT key, value FROM system_settings").fetchall()
    finally:
        conn.close()
    return jsonify({r["key"]: r["value"] for r in rows})


@app.put("/api/settings")
@super_required
def api_update_settings():
    """更新系统设置"""
    d = request.get_json(silent=True) or {}
    conn = get_db()
    try:
        now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        for k, v in d.items():
            exists = conn.execute("SELECT 1 FROM system_settings WHERE key = ?", (k,)).fetchone()
            if exists:
                conn.execute("UPDATE system_settings SET value = ?, updated_at = ? WHERE key = ?", (str(v), now_str, k))
            else:
                conn.execute("INSERT INTO system_settings (key, value, updated_at) VALUES (?, ?, ?)", (k, str(v), now_str))
        log_operation(conn, "修改系统设置", "系统设置", ", ".join(d.keys()))
        conn.commit()
    finally:
        conn.close()
    return jsonify({"message": "设置已保存"})


@app.post("/api/settings/backup")
@super_required
def api_backup():
    """一键备份数据库"""
    try:
        filename = backup_database()
        conn = get_db()
        try:
            log_operation(conn, "数据库备份", "系统设置", filename)
            conn.commit()
        finally:
            conn.close()
        return jsonify({"message": "备份成功", "filename": filename})
    except Exception as e:
        return jsonify({"error": f"备份失败: {str(e)}"}), 500


@app.get("/api/settings/backups")
@super_required
def api_list_backups():
    """获取备份文件列表"""
    return jsonify(list_backups())


@app.delete("/api/settings/backups/<path:filename>")
@super_required
def api_delete_backup(filename):
    """删除备份文件"""
    # 安全检查：只允许删除 backups 目录下的 .db 文件
    if not filename.endswith(".db") or "/" in filename or "\\" in filename:
        return jsonify({"error": "无效的文件名"}), 400
    path = os.path.join(BACKUP_DIR, filename)
    if not os.path.exists(path):
        return jsonify({"error": "备份文件不存在"}), 404
    os.remove(path)
    conn = get_db()
    try:
        log_operation(conn, "删除备份", "系统设置", filename)
        conn.commit()
    finally:
        conn.close()
    return jsonify({"message": "备份已删除"})


@app.get("/api/settings/backups/<path:filename>/download")
@super_required
def api_download_backup(filename):
    """下载备份文件到本地"""
    if not filename.endswith(".db") or "/" in filename or "\\" in filename:
        return jsonify({"error": "无效的文件名"}), 400
    path = os.path.join(BACKUP_DIR, filename)
    if not os.path.exists(path):
        return jsonify({"error": "备份文件不存在"}), 404
    return send_file(path, as_attachment=True, download_name=filename)


@app.post("/api/settings/backups/<path:filename>/restore")
@super_required
def api_restore_backup(filename):
    """恢复备份：用指定备份文件覆盖当前数据库"""
    if not filename.endswith(".db") or "/" in filename or "\\" in filename:
        return jsonify({"error": "无效的文件名"}), 400
    src = os.path.join(BACKUP_DIR, filename)
    if not os.path.exists(src):
        return jsonify({"error": "备份文件不存在"}), 404
    try:
        # 1) 校验备份文件是可用的 SQLite 数据库（防止拿任意文件覆盖导致数据全毁）
        test = sqlite3.connect(src)
        try:
            tables = {r[0] for r in test.execute(
                "SELECT name FROM sqlite_master WHERE type='table'")}
        finally:
            test.close()
        if "students" not in tables or "users" not in tables:
            return jsonify({"error": "备份文件结构异常，不是有效的系统数据库，已取消恢复"}), 400

        # 2) 恢复前先把当前数据库自动备份一份（可一键撤销本次恢复）
        os.makedirs(BACKUP_DIR, exist_ok=True)
        safety_name = f"students_恢复前自动备份_{datetime.now().strftime('%Y%m%d_%H%M%S')}.db"
        shutil.copy2(DB_PATH, os.path.join(BACKUP_DIR, safety_name))

        # 3) 覆盖主库文件。SQLite WAL 模式下运行时直接复制可能复制到 WAL 未落盘内容，
        #    先执行 checkpoint 把 WAL 内容全部写回主文件再复制，恢复才完整。
        conn = get_db()
        try:
            conn.execute("PRAGMA wal_checkpoint(TRUNCATE)")
        except sqlite3.Error:
            pass
        finally:
            conn.close()
        shutil.copy2(src, DB_PATH)
        # 操作日志写入恢复后的新库；日志表异常不应误报恢复失败
        try:
            conn_log = get_db()
            try:
                log_operation(conn_log, "恢复备份", "系统设置", filename,
                              f"恢复前已自动备份为 {safety_name}")
                conn_log.commit()
            finally:
                conn_log.close()
        except sqlite3.Error:
            pass
        return jsonify({"message": f"恢复成功，恢复前的数据已自动备份为 {safety_name}，"
                                   "建议刷新页面并重新登录后确认数据"})
    except sqlite3.Error as e:
        return jsonify({"error": f"恢复失败：备份文件无法读取（{e}）"}), 500
    except Exception as e:
        return jsonify({"error": f"恢复失败: {str(e)}"}), 500


@app.get("/api/settings/operation-logs")
@permission_required("settings.logs", "当前角色无权查看操作日志")
def api_operation_logs():
    """操作日志查询"""
    start = request.args.get("start", "").strip()
    end = request.args.get("end", "").strip()
    username = request.args.get("username", "").strip()
    action = request.args.get("action", "").strip()
    module = request.args.get("module", "").strip()
    where, params = [], []
    if re.match(r"^\d{4}-\d{2}-\d{2}$", start):
        where.append("created_at >= ?")
        params.append(start + " 00:00:00")
    if re.match(r"^\d{4}-\d{2}-\d{2}$", end):
        where.append("created_at <= ?")
        params.append(end + " 23:59:59")
    if username:
        where.append("username LIKE ?")
        params.append(f"%{username}%")
    if action:
        where.append("action LIKE ?")
        params.append(f"%{action}%")
    if module:
        where.append("module LIKE ?")
        params.append(f"%{module}%")
    cond = ("WHERE " + " AND ".join(where)) if where else ""
    conn = get_db()
    try:
        rows = conn.execute(
            f"SELECT * FROM operation_logs {cond} ORDER BY id DESC LIMIT 500",
            params,
        ).fetchall()
    finally:
        conn.close()
    return jsonify([dict(r) for r in rows])


# ---------- v2.0 退费管理 ----------
@app.route("/refunds")
@login_required
def refunds_page():
    """退费管理页"""
    return render_template("refunds.html", user=current_user())


@app.get("/api/refunds")
def list_refunds():
    """退费记录列表，支持日期范围与学生姓名筛选"""
    start = request.args.get("start", "").strip()
    end = request.args.get("end", "").strip()
    keyword = request.args.get("keyword", "").strip()
    where, params = [], []
    if re.match(r"^\d{4}-\d{2}-\d{2}$", start):
        where.append("r.refund_date >= ?")
        params.append(start)
    if re.match(r"^\d{4}-\d{2}-\d{2}$", end):
        where.append("r.refund_date <= ?")
        params.append(end)
    if keyword:
        where.append("(s.name LIKE ? OR s.student_no LIKE ?)")
        params.append(f"%{keyword}%")
        params.append(f"%{keyword}%")
    cond = ("WHERE " + " AND ".join(where)) if where else ""
    conn = get_db()
    try:
        rows = conn.execute(
            f"""
            SELECT r.id, r.student_id, r.payment_id, r.amount, r.lesson_count, r.refund_date,
                   r.method, r.reason, r.remark, r.created_at,
                   s.student_no, s.name AS student_name,
                   p.course AS payment_course, p.amount AS payment_amount, p.lesson_count AS payment_lessons
            FROM refunds r
            JOIN students s ON s.id = r.student_id
            LEFT JOIN payments p ON p.id = r.payment_id
            {cond}
            ORDER BY r.refund_date DESC, r.id DESC
            LIMIT 5000
            """,
            params,
        ).fetchall()
        result = _attach_photos(conn, rows, "refund_photos", "refund_id")
    finally:
        conn.close()
    return jsonify(result)


@app.post("/api/refunds")
@permission_required("refund.write", "当前角色无权新增退费记录")
def create_refund():
    """新建退费记录"""
    d = request.get_json(silent=True) or {}
    student_id = d.get("student_id")
    if not isinstance(student_id, int):
        return jsonify({"error": "请选择学生"}), 400
    amount = _to_float(d.get("amount"))
    if amount is None or amount <= 0:
        return jsonify({"error": "退费金额必须大于 0"}), 400
    payment_id = d.get("payment_id")
    refund_date = str(d.get("refund_date") or "").strip()
    if not re.match(r"^\d{4}-\d{2}-\d{2}$", refund_date):
        refund_date = datetime.now().strftime("%Y-%m-%d")
    method = str(d.get("method") or "").strip() or "原路退回"
    reason = str(d.get("reason") or "").strip()
    remark = str(d.get("remark") or "").strip()[:200]
    lesson_count = d.get("lesson_count")
    if lesson_count in (None, ""):
        lesson_count = 0
    else:
        lesson_count = _to_int(lesson_count)
        if lesson_count is None:
            return jsonify({"error": "退课时数必须为 0 或正整数"}), 400

    conn = get_db()
    try:
        if not conn.execute("SELECT 1 FROM students WHERE id = ?", (student_id,)).fetchone():
            return jsonify({"error": "学生档案不存在"}), 404
        # 如果关联了缴费记录，检查退费金额不超过原缴费
        if payment_id:
            pay = conn.execute("SELECT amount, lesson_count FROM payments WHERE id = ?", (payment_id,)).fetchone()
            if not pay:
                return jsonify({"error": "关联的缴费记录不存在"}), 404
            # 检查已退费总额
            total_refunded = conn.execute(
                "SELECT COALESCE(SUM(amount), 0) AS s FROM refunds WHERE payment_id = ?",
                (payment_id,),
            ).fetchone()["s"]
            if total_refunded + amount > pay["amount"]:
                return jsonify({"error": f"退费金额超过原缴费金额（已退 {total_refunded}，本次退 {amount}，原缴费 {pay['amount']}）"}), 400
            # 检查累计退课时不超过原缴费节数
            total_refunded_lessons = conn.execute(
                "SELECT COALESCE(SUM(lesson_count), 0) AS s FROM refunds WHERE payment_id = ?",
                (payment_id,),
            ).fetchone()["s"]
            orig_lessons = pay["lesson_count"] or 0
            if total_refunded_lessons + lesson_count > orig_lessons:
                return jsonify({"error": f"退课时数超过原缴费节数（已退 {total_refunded_lessons}，本次退 {lesson_count}，原缴费 {orig_lessons} 节）"}), 400
        cur = conn.execute(
            """INSERT INTO refunds (student_id, payment_id, amount, lesson_count, refund_date, method, reason, remark, created_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)""",
            (student_id, payment_id, amount, lesson_count, refund_date, method, reason, remark,
             datetime.now().strftime("%Y-%m-%d %H:%M:%S")),
        )
        log_operation(conn, "创建退费", "退费管理", f"学生ID:{student_id} 金额:{amount} 退课时:{lesson_count}")
        conn.commit()
        rid = cur.lastrowid
    finally:
        conn.close()
    return jsonify({"id": rid, "message": "退费记录已保存"}), 201


@app.put("/api/refunds/<int:refund_id>")
@permission_required("refund.write", "当前角色无权修改退费记录")
def update_refund(refund_id):
    """修改退费记录"""
    d = request.get_json(silent=True) or {}
    amount = _to_float(d.get("amount"))
    if amount is None or amount <= 0:
        return jsonify({"error": "退费金额必须大于 0"}), 400
    refund_date = str(d.get("refund_date") or "").strip()
    if not re.match(r"^\d{4}-\d{2}-\d{2}$", refund_date):
        refund_date = datetime.now().strftime("%Y-%m-%d")
    method = str(d.get("method") or "").strip() or "原路退回"
    reason = str(d.get("reason") or "").strip()
    remark = str(d.get("remark") or "").strip()[:200]
    lesson_count = d.get("lesson_count")
    if lesson_count in (None, ""):
        lesson_count = 0
    else:
        lesson_count = _to_int(lesson_count)
        if lesson_count is None:
            return jsonify({"error": "退课时数必须为 0 或正整数"}), 400
    conn = get_db()
    try:
        old = conn.execute(
            "SELECT student_id, payment_id, lesson_count FROM refunds WHERE id = ?", (refund_id,)
        ).fetchone()
        if not old:
            return jsonify({"error": "退费记录不存在"}), 404
        # 关联缴费时校验累计退课时不超过原缴费节数（排除本条自身旧值）
        if old["payment_id"]:
            pay = conn.execute(
                "SELECT lesson_count FROM payments WHERE id = ?", (old["payment_id"],)
            ).fetchone()
            if pay:
                refunded_lessons = conn.execute(
                    "SELECT COALESCE(SUM(lesson_count), 0) AS s FROM refunds WHERE payment_id = ? AND id != ?",
                    (old["payment_id"], refund_id),
                ).fetchone()["s"]
                orig_lessons = pay["lesson_count"] or 0
                if refunded_lessons + lesson_count > orig_lessons:
                    return jsonify({"error": f"退课时数超过原缴费节数（已退 {refunded_lessons}，本次退 {lesson_count}，原缴费 {orig_lessons} 节）"}), 400
        cur = conn.execute(
            """UPDATE refunds SET amount = ?, lesson_count = ?, refund_date = ?, method = ?, reason = ?, remark = ? WHERE id = ?""",
            (amount, lesson_count, refund_date, method, reason, remark, refund_id),
        )
        if cur.rowcount == 0:
            return jsonify({"error": "退费记录不存在"}), 404
        log_operation(conn, "修改退费", "退费管理", f"退费ID:{refund_id} 退课时:{lesson_count}")
        conn.commit()
    finally:
        conn.close()
    return jsonify({"message": "修改成功"})


@app.delete("/api/refunds/<int:refund_id>")
@permission_required("refund.delete", "教务不可删除退费记录，删除需由管理员或校长操作")
def delete_refund(refund_id):
    """删除退费记录"""
    conn = get_db()
    try:
        cur = conn.execute("DELETE FROM refunds WHERE id = ?", (refund_id,))
        if cur.rowcount == 0:
            return jsonify({"error": "退费记录不存在"}), 404
        conn.execute("DELETE FROM refund_photos WHERE refund_id = ?", (refund_id,))
        log_operation(conn, "删除退费", "退费管理", f"退费ID:{refund_id}")
        conn.commit()
    finally:
        conn.close()
    return jsonify({"message": "删除成功"})


# ---------- 退费证明图片 ----------
MAX_REFUND_PHOTOS = 5   # 每笔退费最多证明图片数


@app.post("/api/refunds/<int:refund_id>/photos")
@permission_required("refund.write", "当前角色无权上传退费凭证")
def upload_refund_photos(refund_id):
    """上传退费证明图片（每笔退费最多 5 张，jpg/png/gif/webp，单张 5MB 内）"""
    files = [f for f in request.files.getlist("photos") if f and f.filename]
    if not files:
        return jsonify({"error": "未选择任何图片"}), 400
    conn = get_db()
    try:
        if not conn.execute(
            "SELECT 1 FROM refunds WHERE id = ?", (refund_id,)
        ).fetchone():
            return jsonify({"error": "退费记录不存在"}), 404
        existing = conn.execute(
            "SELECT COUNT(*) AS c FROM refund_photos WHERE refund_id = ?", (refund_id,)
        ).fetchone()["c"]
        if existing + len(files) > MAX_REFUND_PHOTOS:
            return jsonify(
                {"error": f"退费证明最多 {MAX_REFUND_PHOTOS} 张（已有 {existing} 张，本次选择 {len(files)} 张）"}
            ), 400

        # 先整体校验（格式/大小），全部通过后再统一保存
        blobs = []
        for f in files:
            ext = os.path.splitext(f.filename)[1].lower()
            if ext not in ALLOWED_PHOTO_EXT:
                return jsonify(
                    {"error": f"不支持的图片格式：{f.filename}（仅支持 jpg/png/gif/webp）"}
                ), 400
            blob = f.read()
            if len(blob) > MAX_PHOTO_SIZE:
                return jsonify(
                    {"error": f"图片过大：{f.filename}（单张不超过 5MB）"}
                ), 400
            blobs.append((blob, ext, os.path.basename(f.filename).replace(" ", "_") or "photo"))

        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        for blob, ext, original in blobs:
            stored = f"rf{refund_id}_{uuid.uuid4().hex[:12]}{ext}"
            conn.execute(
                "INSERT INTO refund_photos (refund_id, filename, original_name, size, uploaded_at, data) "
                "VALUES (?, ?, ?, ?, ?, ?)",
                (refund_id, stored, original, len(blob), now, blob),
            )
        conn.commit()
    finally:
        conn.close()
    return jsonify({"message": f"成功上传 {len(blobs)} 张退费证明"}), 201


@app.delete("/api/refund-photos/<int:photo_id>")
@permission_required("refund.delete", "当前角色无权删除退费凭证")
def delete_refund_photo(photo_id):
    """删除一张退费证明图片"""
    conn = get_db()
    try:
        cur = conn.execute("DELETE FROM refund_photos WHERE id = ?", (photo_id,))
        conn.commit()
        if cur.rowcount == 0:
            return jsonify({"error": "图片不存在"}), 404
    finally:
        conn.close()
    return jsonify({"message": "删除成功"})


@app.get("/uploads/refund-photos/<path:filename>")
def get_refund_photo(filename):
    """查看退费证明图片；带 ?download=1 时以原始文件名下载（只读角色不可下载）"""
    if _viewer_download_denied():
        return jsonify({"error": "只读角色无权下载退费证明"}), 403
    conn = get_db()
    try:
        row = conn.execute(
            "SELECT data, original_name FROM refund_photos WHERE filename = ?", (filename,)
        ).fetchone()
    finally:
        conn.close()
    if not row or not row["data"]:
        abort(404)
    download_name = row["original_name"] or filename
    return send_file(
        io.BytesIO(row["data"]),
        download_name=download_name,
        as_attachment=bool(request.args.get("download")),
    )


@app.get("/api/refunds/export")
@permission_required("refund.export", "当前角色无权导出退费 Excel")
def export_refunds():
    """导出退费查询结果为 Excel（筛选条件与 /api/refunds 一致）"""
    start = request.args.get("start", "").strip()
    end = request.args.get("end", "").strip()
    keyword = request.args.get("keyword", "").strip()
    where, params = [], []
    if re.match(r"^\d{4}-\d{2}-\d{2}$", start):
        where.append("r.refund_date >= ?")
        params.append(start)
    if re.match(r"^\d{4}-\d{2}-\d{2}$", end):
        where.append("r.refund_date <= ?")
        params.append(end)
    if keyword:
        where.append("(s.name LIKE ? OR s.student_no LIKE ?)")
        params.append(f"%{keyword}%")
        params.append(f"%{keyword}%")
    cond = ("WHERE " + " AND ".join(where)) if where else ""

    conn = get_db()
    try:
        rows = conn.execute(
            f"""
            SELECT r.refund_date, s.student_no, s.name AS student_name,
                   p.course AS payment_course, p.amount AS payment_amount, p.lesson_count AS payment_lessons,
                   r.amount, r.lesson_count, r.method, r.reason, r.remark
            FROM refunds r
            JOIN students s ON s.id = r.student_id
            LEFT JOIN payments p ON p.id = r.payment_id
            {cond}
            ORDER BY r.refund_date DESC, r.id DESC
            LIMIT 5000
            """,
            params,
        ).fetchall()
    finally:
        conn.close()

    # ---- 生成 Excel ----
    wb = Workbook()
    ws = wb.active
    ws.title = "退费查询"

    head_font = Font(name="等线", bold=True, color="FFFFFF")
    head_fill = PatternFill("solid", fgColor="DC2626")
    cell_font = Font(name="等线")
    money_font = Font(name="等线", color="C00000")
    center = Alignment(horizontal="center", vertical="center")
    left = Alignment(horizontal="left", vertical="center", wrap_text=True)
    border = Border(
        left=Side(style="thin", color="D0D0D0"),
        right=Side(style="thin", color="D0D0D0"),
        top=Side(style="thin", color="D0D0D0"),
        bottom=Side(style="thin", color="D0D0D0"),
    )

    # 标题行
    ws.merge_cells("A1:K1")
    ws["A1"] = "退费记录查询"
    ws["A1"].font = Font(name="等线", size=14, bold=True)
    ws["A1"].alignment = center
    ws.row_dimensions[1].height = 32

    # 表头
    headers = ["退费日期", "学号", "学生姓名", "关联课程", "原缴费金额",
               "原缴费节数", "退费金额", "退课时数", "退费方式", "退费原因", "备注"]
    ws.append(headers)
    for col_num, _ in enumerate(headers, 1):
        c = ws.cell(row=2, column=col_num)
        c.font = head_font
        c.fill = head_fill
        c.alignment = center
        c.border = border
    ws.row_dimensions[2].height = 26

    # 数据行
    for r in rows:
        ws.append([
            r["refund_date"] or "",
            r["student_no"] or "",
            r["student_name"] or "",
            r["payment_course"] or "",
            r["payment_amount"] or 0,
            r["payment_lessons"] or 0,
            r["amount"] or 0,
            r["lesson_count"] or 0,
            r["method"] or "",
            r["reason"] or "",
            r["remark"] or "",
        ])

    # 样式 + 列宽
    for row in ws.iter_rows(min_row=3, max_row=ws.max_row):
        for cell in row:
            cell.font = money_font if cell.column in (5, 7) else cell_font
            cell.alignment = left if cell.column in (4, 10, 11) else center
            cell.border = border
        ws.row_dimensions[row[0].row].height = 22

    for col, width in zip("ABCDEFGHIJK", (14, 14, 12, 18, 14, 12, 14, 12, 14, 18, 24)):
        ws.column_dimensions[col].width = width

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    end_str = end or datetime.now().strftime("%Y%m%d")
    return send_file(
        buf,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name=f"退费记录_{end_str}.xlsx",
    )


# ---------- v2.0 数据导入 ----------
@app.route("/import")
@permission_required("import.all", "当前角色无权使用数据导入")
def import_page():
    """数据导入页"""
    return render_template("import.html", user=current_user())


@app.get("/api/import/template")
def import_template():
    """下载学生导入模板（Excel，含初始购课课时列）"""
    wb = Workbook()
    ws = wb.active
    ws.title = "学生导入模板"
    headers = ["姓名", "性别", "出生日期", "身份证号", "电话", "学校", "年级", "课程",
               "已购课时", "缴费金额(元)", "缴费日期", "地址", "备注",
               "监护人姓名", "监护人电话", "监护人微信"]
    widths = [12, 8, 14, 20, 15, 20, 10, 16, 10, 12, 14, 30, 30, 12, 15, 15]
    head_font = Font(name="等线", bold=True, color="FFFFFF")
    head_fill = PatternFill("solid", fgColor="4472C4")
    center = Alignment(horizontal="center", vertical="center")
    thin = Side(style="thin", color="D9D9D9")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for i, (h, w) in enumerate(zip(headers, widths), start=1):
        c = ws.cell(row=1, column=i, value=h)
        c.font = head_font
        c.fill = head_fill
        c.alignment = center
        c.border = border
        ws.column_dimensions[get_column_letter(i)].width = w
    # 示例行
    examples = ["张三", "男", "2010-05-15", "440105201005150011", "13800138000",
                "实验中学", "七年级", "英语", "50", "6000", "2026-09-17",
                "XX路XX号", "优秀学生", "张父", "13800138001", "zhangfather"]
    for i, v in enumerate(examples, start=1):
        c = ws.cell(row=2, column=i, value=v)
        c.border = border
        c.font = Font(name="等线", color="808080")
    # 说明行：课时与剩余课时的口径
    note = ("说明：填写「已购课时」或「缴费金额」后，导入时自动生成一条缴费记录，"
            "剩余课时 = 已购课时 − 已退课时 − 已消耗课时（新学员无退费与消耗，剩余课时即已购课时）；两列留空则不生成缴费记录")
    nc = ws.cell(row=3, column=1, value=note)
    nc.font = Font(name="等线", size=9, color="808080")
    ws.merge_cells(start_row=3, start_column=1, end_row=3, end_column=16)
    ws.row_dimensions[1].height = 24
    ws.freeze_panes = "A2"
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return send_file(
        buf,
        as_attachment=True,
        download_name="学生导入模板.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.post("/api/import/students")
@permission_required("import.all", "当前角色无权使用数据导入")
def import_students():
    """批量导入学生（Excel）"""
    from openpyxl import load_workbook
    f = request.files.get("file")
    if not f or not f.filename:
        return jsonify({"error": "请选择 Excel 文件"}), 400
    if not f.filename.endswith((".xlsx", ".xls")):
        return jsonify({"error": "仅支持 .xlsx 格式文件"}), 400
    try:
        wb = load_workbook(f, read_only=True, data_only=True)
    except Exception:
        return jsonify({"error": "Excel 文件解析失败"}), 400
    ws = wb.active
    rows = list(ws.iter_rows(min_row=2, values_only=True))
    if not rows:
        return jsonify({"error": "Excel 中没有数据行"}), 400
    if len(rows) > 500:
        return jsonify({"error": "单次导入不超过 500 条"}), 400

    # 按表头名定位列（兼容列序不同的旧模板；第 1 行是表头，数据从第 2 行起）
    header_row = list(ws.iter_rows(min_row=1, max_row=1, values_only=True))[0]
    header_pos = {}
    for ci, h in enumerate(header_row):
        h = str(h or "").strip()
        if h:
            header_pos[h] = ci
    # 旧模板固定列序兜底：无表头时按原 13 列顺序解析
    col = {
        "name": header_pos.get("姓名", 0), "gender": header_pos.get("性别", 1),
        "birthdate": header_pos.get("出生日期", 2), "id_card": header_pos.get("身份证号", 3),
        "phone": header_pos.get("电话", 4), "school": header_pos.get("学校", 5),
        "grade": header_pos.get("年级", 6), "course": header_pos.get("课程", 7),
        "lessons": header_pos.get("已购课时"), "amount": header_pos.get("缴费金额(元)"),
        "pay_date": header_pos.get("缴费日期"), "address": header_pos.get("地址", 8),
        "notes": header_pos.get("备注", 9),
        "g_name": header_pos.get("监护人姓名", 10), "g_phone": header_pos.get("监护人电话", 11),
        "g_wechat": header_pos.get("监护人微信", 12),
    }

    def cell(row, key):
        ci = col[key]
        if ci is None or ci >= len(row):
            return ""
        v = row[ci]
        # openpyxl 日期单元格转字符串
        if hasattr(v, "strftime"):
            try:
                return v.strftime("%Y-%m-%d")
            except Exception:
                return str(v).strip()
        return str(v or "").strip() if v is not None else ""

    def to_num(v):
        v = (v or "").replace(",", "").replace("，", "").strip()
        if not v:
            return None
        try:
            return float(v)
        except ValueError:
            return None

    conn = get_db()
    success, fail = 0, 0
    errors = []
    try:
        for idx, row in enumerate(rows, start=2):
            name = cell(row, "name")
            if not name:
                fail += 1
                errors.append(f"第{idx}行：姓名为空，跳过")
                continue
            gender = cell(row, "gender") or "男"
            if gender not in ("男", "女"):
                gender = "男"
            birthdate = cell(row, "birthdate")
            id_card = cell(row, "id_card")
            phone = cell(row, "phone")
            school = cell(row, "school")
            grade = cell(row, "grade")
            course = cell(row, "course")
            address = cell(row, "address")
            notes = cell(row, "notes")
            g_name = cell(row, "g_name")
            g_phone = cell(row, "g_phone")
            g_wechat = cell(row, "g_wechat")
            # 初始购课信息（可为空）
            lessons = to_num(cell(row, "lessons"))
            amount = to_num(cell(row, "amount"))
            pay_date = cell(row, "pay_date")
            if lessons is not None and lessons < 0:
                lessons = None
            if amount is not None and amount < 0:
                amount = None
            if pay_date and not re.match(r"^\d{4}-\d{2}-\d{2}$", pay_date):
                pay_date = ""  # 格式不对时按导入当天处理

            try:
                student_no = next_student_no(conn)
            except ValueError:
                fail += 1
                errors.append(f"第{idx}行：学生编号已用尽")
                continue
            now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            conn.execute(
                """INSERT INTO students (student_no, name, gender, birthdate, id_card, phone,
                   school, grade, course, address, notes, created_at, updated_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                (student_no, name, gender, birthdate, id_card, phone,
                 school, grade, course, address, notes, now_str, now_str),
            )
            sid = conn.execute("SELECT last_insert_rowid()").fetchone()[0]
            if g_name or g_phone or g_wechat:
                conn.execute(
                    "INSERT INTO guardians (student_id, name, phone, wechat) VALUES (?, ?, ?, ?)",
                    (sid, g_name, g_phone, g_wechat),
                )
            # 填了已购课时或缴费金额时，自动生成一条初始缴费记录（剩余课时即由此而来）
            if lessons is not None or amount is not None:
                lc = int(lessons) if lessons is not None else 0
                amt = amount if amount is not None else 0.0
                conn.execute(
                    """
                    INSERT INTO payments
                        (student_id, course, amount, unit_price, lesson_count,
                         discount_percent, discount_amount, pay_date, method,
                         remark, status, created_at)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (sid, course or "", amt, 0, lc, 100, 0,
                     pay_date or datetime.now().strftime("%Y-%m-%d"),
                     "现金", "批量导入自动生成", "已确认",
                     datetime.now().strftime("%Y-%m-%d %H:%M:%S")),
                )
            success += 1
        log_operation(conn, "批量导入学生", "数据导入", f"成功:{success} 失败:{fail}")
        conn.commit()
    finally:
        conn.close()
    return jsonify({
        "message": f"导入完成：成功 {success} 条，失败 {fail} 条",
        "success": success,
        "fail": fail,
        "errors": errors[:20],
    })


# ---------- 页面 ----------
@app.route("/")
@login_required
def home():
    """系统首页：模块导航与概况统计"""
    conn = get_db()
    try:
        u = current_user(conn)
        stats = {
            "total": conn.execute("SELECT COUNT(*) AS c FROM students").fetchone()["c"],
            "male": conn.execute(
                "SELECT COUNT(*) AS c FROM students WHERE gender = '男'").fetchone()["c"],
            "female": conn.execute(
                "SELECT COUNT(*) AS c FROM students WHERE gender = '女'").fetchone()["c"],
        }
        month_prefix = datetime.now().strftime("%Y-%m")
        stats["month_new"] = conn.execute(
            "SELECT COUNT(*) AS c FROM students WHERE created_at LIKE ?",
            (month_prefix + "%",),
        ).fetchone()["c"]
        stats["photos"] = conn.execute("SELECT COUNT(*) AS c FROM photos").fetchone()["c"]
        stats["attachments"] = conn.execute(
            "SELECT COUNT(*) AS c FROM attachments").fetchone()["c"]
        stats["attendance"] = conn.execute(
            "SELECT COUNT(*) AS c FROM attendance").fetchone()["c"]
        stats["courses"] = conn.execute(
            "SELECT COUNT(*) AS c FROM courses").fetchone()["c"]
        stats["subjects"] = conn.execute(
            "SELECT COUNT(*) AS c FROM subjects").fetchone()["c"]
        stats["teachers"] = conn.execute(
            "SELECT COUNT(*) AS c FROM teachers").fetchone()["c"]
        stats["payment_total"] = conn.execute(
            "SELECT COALESCE(SUM(amount), 0) AS s FROM payments").fetchone()["s"]
        stats["payment_month"] = conn.execute(
            "SELECT COALESCE(SUM(amount), 0) AS s FROM payments WHERE pay_date LIKE ?",
            (month_prefix + "%",),
        ).fetchone()["s"]
        stats["payment_today"] = conn.execute(
            "SELECT COALESCE(SUM(amount), 0) AS s FROM payments WHERE pay_date = ?",
            (datetime.now().strftime("%Y-%m-%d"),),
        ).fetchone()["s"]
        # 课时统计（与学生列表口径一致：剩余 = 已购 - 已退 - 已消耗）
        lesson_rows = _lesson_stats_rows()
        stats["lessons_remaining"] = sum(r["remaining_lessons"] for r in lesson_rows)
        stats["lessons_warn"] = sum(1 for r in lesson_rows if r["warn"])
        recent = [dict(r) for r in conn.execute(
            """
            SELECT id, student_no, name, gender, course, created_at
            FROM students ORDER BY id DESC LIMIT 6
            """
        ).fetchall()]
    finally:
        conn.close()
    return render_template("home.html", stats=stats, recent=recent, host=request.host, user=u)


@app.route("/students")
@login_required
def index():
    """学生档案列表页"""
    user = current_user()
    role = (user or {}).get("role")
    return render_template(
        "index.html", user=user,
        can_create=can_create_student(role),      # 新建档案（教师可）
        can_edit=can_write_student(role),          # 编辑 / 删除档案（仅管理员 / 校长 / 教务）
        can_edit_resume=can_write_resume(role),    # 编辑简历（含教师）
    )


# ---------- 接口 ----------
@app.get("/api/students")
def list_students():
    """学生列表，支持关键字搜索（姓名/电话/身份证/监护人/课程/学校）+ 分页"""
    keyword = request.args.get("keyword", "").strip()
    page = max(1, request.args.get("page", 1, type=int))
    per_page = request.args.get("per_page", 20, type=int)
    per_page = max(1, min(per_page, 200))
    conn = get_db()
    try:
        if keyword:
            like = f"%{keyword}%"
            where_sql = """
                WHERE name LIKE ? OR student_no LIKE ? OR phone LIKE ? OR id_card LIKE ?
                   OR course LIKE ? OR school LIKE ?
                   OR id IN (SELECT student_id FROM guardians
                             WHERE name LIKE ? OR phone LIKE ? OR wechat LIKE ?)
                """
            params = (like, like, like, like, like, like, like, like, like)
        else:
            where_sql = ""
            params = ()

        total = conn.execute(
            f"SELECT COUNT(*) AS c FROM students {where_sql}", params
        ).fetchone()["c"]

        # 统计性别与本月建档（基于全部搜索结果，非仅当前页）
        month_prefix = datetime.now().strftime("%Y-%m")
        stats = {
            "total": total,
            "male": conn.execute(
                f"SELECT COUNT(*) AS c FROM students {where_sql} AND gender = '男'", params
            ).fetchone()["c"] if where_sql else conn.execute(
                "SELECT COUNT(*) AS c FROM students WHERE gender = '男'"
            ).fetchone()["c"],
            "female": conn.execute(
                f"SELECT COUNT(*) AS c FROM students {where_sql} AND gender = '女'", params
            ).fetchone()["c"] if where_sql else conn.execute(
                "SELECT COUNT(*) AS c FROM students WHERE gender = '女'"
            ).fetchone()["c"],
            "month_new": conn.execute(
                f"SELECT COUNT(*) AS c FROM students {where_sql} AND created_at LIKE ?", params + (month_prefix + "%",)
            ).fetchone()["c"] if where_sql else conn.execute(
                "SELECT COUNT(*) AS c FROM students WHERE created_at LIKE ?", (month_prefix + "%",)
            ).fetchone()["c"],
        }

        rows = conn.execute(
            f"SELECT * FROM students {where_sql} ORDER BY id DESC LIMIT ? OFFSET ?",
            params + (per_page, (page - 1) * per_page),
        ).fetchall()

        student_ids = [r["id"] for r in rows]
        guardians = {}
        if student_ids:
            placeholders = ",".join("?" * len(student_ids))
            for g in conn.execute(
                f"SELECT * FROM guardians WHERE student_id IN ({placeholders}) ORDER BY student_id, id",
                student_ids,
            ).fetchall():
                guardians.setdefault(g["student_id"], []).append(dict(g))
        photos = {}
        if student_ids:
            for p in conn.execute(
                f"SELECT id, student_id, filename, original_name, uploaded_at, sort FROM photos WHERE student_id IN ({placeholders}) ORDER BY student_id, sort, id",
                student_ids,
            ).fetchall():
                photos.setdefault(p["student_id"], []).append(dict(p))
        # 课时统计：已购课时 = 全部缴费记录购买节数合计（无缴费默认 0）
        lesson_totals = {}
        if student_ids:
            for r in conn.execute(
                f"SELECT student_id, COALESCE(SUM(lesson_count), 0) AS t FROM payments WHERE student_id IN ({placeholders}) GROUP BY student_id",
                student_ids,
            ).fetchall():
                lesson_totals[r["student_id"]] = r["t"]
        # 已退课时 = 全部退费记录退课时数合计（无退费默认 0）
        refunded_lessons = {}
        if student_ids:
            for r in conn.execute(
                f"SELECT student_id, COALESCE(SUM(lesson_count), 0) AS t FROM refunds WHERE student_id IN ({placeholders}) GROUP BY student_id",
                student_ids,
            ).fetchall():
                refunded_lessons[r["student_id"]] = r["t"]
        # 已消耗课时 = 出勤 + 迟到的考勤记录数（与考勤导出口径一致；请假/缺勤不扣课时）
        used_lessons = {}
        if student_ids:
            for r in conn.execute(
                f"SELECT student_id, COUNT(*) AS c FROM attendance WHERE student_id IN ({placeholders}) AND status IN ('present', 'late') GROUP BY student_id",
                student_ids,
            ).fetchall():
                used_lessons[r["student_id"]] = r["c"]
    finally:
        conn.close()

    students = []
    for r in rows:
        s = dict(r)
        s.pop("guardian_name", None)
        s.pop("guardian_phone", None)
        s["guardians"] = guardians.get(s["id"], [])
        s["photos"] = photos.get(s["id"], [])
        s["total_lessons"] = lesson_totals.get(s["id"], 0)
        s["refunded_lessons"] = refunded_lessons.get(s["id"], 0)
        s["used_lessons"] = used_lessons.get(s["id"], 0)
        # 剩余课时 = 已购课时 - 已退课时 - 已消耗课时（可为负，提示课时超上）
        s["remaining_lessons"] = s["total_lessons"] - s["refunded_lessons"] - s["used_lessons"]
        students.append(s)
    return jsonify({"students": students, "total": total, "page": page, "per_page": per_page, "stats": stats})


@app.post("/api/students")
@student_create_required
def create_student():
    """建立学生档案"""
    data = request.get_json(silent=True) or {}
    payload = clean_payload(data)
    guardians = extract_guardians(data)
    error = validate(payload)
    if error:
        return jsonify({"error": error}), 400

    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    conn = get_db()
    try:
        try:
            student_no = next_student_no(conn)
        except ValueError:
            return jsonify({"error": "学生编号已用尽（a00001 - z99999）"}), 400
        cur = conn.execute(
            """
            INSERT INTO students
                (student_no, name, gender, birthdate, id_card, phone,
                 school, grade, course, address, notes, created_at, updated_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                student_no,
                payload["name"], payload["gender"], payload["birthdate"],
                payload["id_card"], payload["phone"],
                payload["school"], payload["grade"], payload["course"],
                payload["address"], payload["notes"], now, now,
            ),
        )
        student_id = cur.lastrowid
        for g in guardians:
            conn.execute(
                "INSERT INTO guardians (student_id, name, phone, wechat) VALUES (?, ?, ?, ?)",
                (student_id, g["name"], g["phone"], g["wechat"]),
            )
        conn.commit()
    finally:
        conn.close()
    return jsonify({"id": student_id, "student_no": student_no, "message": "建档成功"}), 201


@app.put("/api/students/<int:student_id>")
@student_write_required
def update_student(student_id):
    """修改学生档案"""
    data = request.get_json(silent=True) or {}
    payload = clean_payload(data)
    guardians = extract_guardians(data)

    conn = get_db()
    try:
        # 未提交 id_card / notes 的入口（如简历编辑页）保留档案中已有值，避免被清空
        if "id_card" not in data or "notes" not in data:
            exist = conn.execute(
                "SELECT id_card, notes FROM students WHERE id = ?", (student_id,)
            ).fetchone()
            if exist:
                if "id_card" not in data:
                    payload["id_card"] = exist["id_card"] or ""
                if "notes" not in data:
                    payload["notes"] = exist["notes"] or ""

        error = validate(payload)
        if error:
            return jsonify({"error": error}), 400

        cur = conn.execute(
            """
            UPDATE students SET
                name = ?, gender = ?, birthdate = ?, id_card = ?, phone = ?,
                school = ?, grade = ?, course = ?, address = ?, notes = ?,
                updated_at = ?
            WHERE id = ?
            """,
            (
                payload["name"], payload["gender"], payload["birthdate"],
                payload["id_card"], payload["phone"],
                payload["school"], payload["grade"], payload["course"],
                payload["address"], payload["notes"],
                datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                student_id,
            ),
        )
        if cur.rowcount == 0:
            return jsonify({"error": "学生档案不存在"}), 404

        # 监护人：整体替换（先删后插）
        conn.execute("DELETE FROM guardians WHERE student_id = ?", (student_id,))
        for g in guardians:
            conn.execute(
                "INSERT INTO guardians (student_id, name, phone, wechat) VALUES (?, ?, ?, ?)",
                (student_id, g["name"], g["phone"], g["wechat"]),
            )
        conn.commit()
    finally:
        conn.close()
    return jsonify({"message": "修改成功"})


@app.delete("/api/students/<int:student_id>")
@student_write_required
def delete_student(student_id):
    """删除学生档案"""
    conn = get_db()
    try:
        cur = conn.execute(
            "DELETE FROM students WHERE id = ?", (student_id,)
        )
        if cur.rowcount == 0:
            return jsonify({"error": "学生档案不存在"}), 404

        conn.execute("DELETE FROM photos WHERE student_id = ?", (student_id,))
        conn.execute("DELETE FROM attachments WHERE student_id = ?", (student_id,))
        conn.execute("DELETE FROM resume_items WHERE student_id = ?", (student_id,))
        conn.execute("DELETE FROM attendance WHERE student_id = ?", (student_id,))
        conn.execute(
            "DELETE FROM payment_photos WHERE payment_id IN "
            "(SELECT id FROM payments WHERE student_id = ?)",
            (student_id,),
        )
        conn.execute("DELETE FROM payments WHERE student_id = ?", (student_id,))
        conn.execute(
            "DELETE FROM refund_photos WHERE refund_id IN "
            "(SELECT id FROM refunds WHERE student_id = ?)",
            (student_id,),
        )
        conn.execute("DELETE FROM refunds WHERE student_id = ?", (student_id,))
        conn.execute("DELETE FROM enrollments WHERE student_id = ?", (student_id,))
        conn.execute(
            "DELETE FROM guardians WHERE student_id = ?", (student_id,)
        )
        conn.commit()
    finally:
        conn.close()
    return jsonify({"message": "删除成功"})


# ---------- 照片 ----------
@app.post("/api/students/<int:student_id>/photos")
@resume_write_required
def upload_photos(student_id):
    """上传学生照片（最多 5 张）"""
    files = [f for f in request.files.getlist("photos") if f and f.filename]
    if not files:
        return jsonify({"error": "未选择任何照片"}), 400

    conn = get_db()
    try:
        if not conn.execute(
            "SELECT 1 FROM students WHERE id = ?", (student_id,)
        ).fetchone():
            return jsonify({"error": "学生档案不存在"}), 404
        existing = conn.execute(
            "SELECT COUNT(*) AS c FROM photos WHERE student_id = ?", (student_id,)
        ).fetchone()["c"]
        if existing + len(files) > MAX_PHOTOS:
            return jsonify(
                {"error": f"照片最多 {MAX_PHOTOS} 张（已有 {existing} 张，本次选择 {len(files)} 张）"}
            ), 400

        # 先整体校验（格式/大小），全部通过后再统一保存
        blobs = []
        for f in files:
            ext = os.path.splitext(f.filename)[1].lower()
            if ext not in ALLOWED_PHOTO_EXT:
                return jsonify(
                    {"error": f"不支持的图片格式：{f.filename}（仅支持 jpg/png/gif/webp）"}
                ), 400
            blob = f.read()
            if len(blob) > MAX_PHOTO_SIZE:
                return jsonify(
                    {"error": f"图片过大：{f.filename}（单张不超过 5MB）"}
                ), 400
            blobs.append((blob, ext, os.path.basename(f.filename).replace(" ", "_") or "photo"))

        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        for blob, ext, original in blobs:
            stored = f"{student_id}_{uuid.uuid4().hex[:12]}{ext}"
            conn.execute(
                "INSERT INTO photos (student_id, filename, original_name, uploaded_at, data) VALUES (?, ?, ?, ?, ?)",
                (student_id, stored, original, now, blob),
            )
        conn.commit()
    finally:
        conn.close()
    return jsonify({"message": f"成功上传 {len(blobs)} 张照片"}), 201


@app.delete("/api/students/<int:student_id>/photos/<int:photo_id>")
@resume_write_required
def delete_photo(student_id, photo_id):
    """删除一张学生照片"""
    conn = get_db()
    try:
        row = conn.execute(
            "SELECT id, student_id, filename, original_name, uploaded_at, sort FROM photos WHERE id = ? AND student_id = ?", (photo_id, student_id)
        ).fetchone()
        if not row:
            return jsonify({"error": "照片不存在"}), 404
        conn.execute("DELETE FROM photos WHERE id = ?", (photo_id,))
        conn.commit()
    finally:
        conn.close()
    return jsonify({"message": "删除成功"})


@app.get("/api/students/<int:student_id>/photos")
def list_student_photos(student_id):
    """获取学生的照片列表（第一张即证件照）"""
    conn = get_db()
    try:
        if not conn.execute(
            "SELECT 1 FROM students WHERE id = ?", (student_id,)
        ).fetchone():
            return jsonify({"error": "学生档案不存在"}), 404
        rows = conn.execute(
            "SELECT id, student_id, filename, original_name, uploaded_at, sort FROM photos WHERE student_id = ? ORDER BY sort, id", (student_id,)
        ).fetchall()
    finally:
        conn.close()
    return jsonify([dict(r) for r in rows])


@app.post("/api/students/<int:student_id>/photos/<int:photo_id>/set-avatar")
@resume_write_required
def set_avatar_photo(student_id, photo_id):
    """将某张照片设为证件照（排在第一位）"""
    conn = get_db()
    try:
        if not conn.execute(
            "SELECT 1 FROM photos WHERE id = ? AND student_id = ?", (photo_id, student_id)
        ).fetchone():
            return jsonify({"error": "照片不存在"}), 404
        conn.execute("UPDATE photos SET sort = 0 WHERE student_id = ?", (student_id,))
        conn.execute("UPDATE photos SET sort = -1 WHERE id = ?", (photo_id,))
        conn.commit()
    finally:
        conn.close()
    return jsonify({"message": "已设为证件照"})


@app.get("/uploads/photos/<path:filename>")
def get_photo(filename):
    """查看照片；带 ?download=1 时以原始文件名下载（只读角色不可下载）"""
    if _viewer_download_denied():
        return jsonify({"error": "只读角色无权下载学生照片"}), 403
    conn = get_db()
    try:
        row = conn.execute(
            "SELECT data, original_name FROM photos WHERE filename = ?", (filename,)
        ).fetchone()
    finally:
        conn.close()
    if not row or not row["data"]:
        abort(404)
    download_name = row["original_name"] or filename
    return send_file(
        io.BytesIO(row["data"]),
        download_name=download_name,
        as_attachment=bool(request.args.get("download")),
    )


# ---------- 学生简历页 ----------
@app.route("/resume/<int:student_id>")
@login_required
def resume_page(student_id):
    """学生简历（独立页面，可打印/保存为 PDF）"""
    conn = get_db()
    try:
        row = conn.execute(
            "SELECT * FROM students WHERE id = ?", (student_id,)
        ).fetchone()
        if not row:
            abort(404)
        student = dict(row)
        student.pop("guardian_name", None)
        student.pop("guardian_phone", None)
        guardians = [dict(g) for g in conn.execute(
            "SELECT name, phone, wechat FROM guardians WHERE student_id = ? ORDER BY id",
            (student_id,),
        ).fetchall()]
        photos = [dict(p) for p in conn.execute(
            "SELECT id, filename, original_name FROM photos WHERE student_id = ? ORDER BY id",
            (student_id,),
        ).fetchall()]
        sections = load_resume_sections(conn, student_id)
    finally:
        conn.close()

    # 年龄：优先出生日期，缺失时从身份证号推算（仅用于计算，不在简历中展示身份证号）
    age = None
    bd = student.get("birthdate") or ""
    idc = student.get("id_card") or ""
    if not bd and ID_CARD_RE.match(idc):
        bd = f"{idc[6:10]}-{idc[10:12]}-{idc[12:14]}"
    if bd:
        try:
            b = datetime.strptime(bd, "%Y-%m-%d")
            today = datetime.now()
            age = today.year - b.year - ((today.month, today.day) < (b.month, b.day))
        except ValueError:
            pass

    # 「兴趣爱好 / 性格描述」为纯文字项，各取首条内容
    hobby_text = next((it["description"] for it in sections.get("hobby", []) if it.get("description")), "")
    personality_text = next((it["description"] for it in sections.get("personality", []) if it.get("description")), "")

    # 权限：编辑简历（管理员/校长/教务/教师）；导出 PDF（除只读外的所有角色）
    role = (current_user() or {}).get("role")
    return render_template(
        "resume.html",
        student=student, guardians=guardians, photos=photos, sections=sections,
        hobby_text=hobby_text, personality_text=personality_text,
        age=age, user=current_user(),
        can_edit_resume=can_write_resume(role),
        can_export_resume=role != "viewer",
    )


@app.route("/resume/<int:student_id>/edit")
@login_required
@resume_write_required
def resume_edit_page(student_id):
    """简历编辑页（单独修改简历中的培训/项目/证书/获奖等内容）"""
    conn = get_db()
    try:
        row = conn.execute(
            "SELECT * FROM students WHERE id = ?", (student_id,)
        ).fetchone()
        if not row:
            abort(404)
        student = dict(row)
        student.pop("guardian_name", None)
        student.pop("guardian_phone", None)
        guardians = [dict(g) for g in conn.execute(
            "SELECT name, phone, wechat FROM guardians WHERE student_id = ? ORDER BY id",
            (student_id,),
        ).fetchall()]
    finally:
        conn.close()

    # 年龄：优先出生日期，缺失时从身份证号推算（与简历页一致）
    age = None
    bd = student.get("birthdate") or ""
    idc = student.get("id_card") or ""
    if not bd and ID_CARD_RE.match(idc):
        bd = f"{idc[6:10]}-{idc[10:12]}-{idc[12:14]}"
    if bd:
        try:
            b = datetime.strptime(bd, "%Y-%m-%d")
            today = datetime.now()
            age = today.year - b.year - ((today.month, today.day) < (b.month, b.day))
        except ValueError:
            pass

    conn = get_db()
    try:
        sections = load_resume_sections(conn, student_id)
    finally:
        conn.close()

    role = (current_user() or {}).get("role")
    return render_template(
        "resume_edit.html",
        student=student, guardians=guardians, age=age, user=current_user(),
        can_edit_profile=can_write_student(role),   # 档案信息区：仅管理员 / 校长 / 教务可改
        text_initial={
            "hobby": next((it["description"] for it in sections.get("hobby", []) if it.get("description")), ""),
            "personality": next((it["description"] for it in sections.get("personality", []) if it.get("description")), ""),
            "self_evaluation": student.get("notes") or "",
        },
    )


@app.put("/api/students/<int:student_id>/notes")
@resume_write_required
def update_notes(student_id):
    """修改自我评价（对应档案备注字段，简历中显示为自我评价）"""
    d = request.get_json(silent=True) or {}
    notes = str(d.get("notes") or "").strip()[:500]
    conn = get_db()
    try:
        cur = conn.execute(
            "UPDATE students SET notes = ?, updated_at = ? WHERE id = ?",
            (notes, datetime.now().strftime("%Y-%m-%d %H:%M:%S"), student_id),
        )
        conn.commit()
        if cur.rowcount == 0:
            return jsonify({"error": "学生档案不存在"}), 404
    finally:
        conn.close()
    return jsonify({"message": "保存成功"})


@app.put("/api/students/<int:student_id>/resume-text")
@resume_write_required
def save_resume_text(student_id):
    """保存简历中的纯文字项（兴趣爱好 / 性格描述）：每个类型只保留一条，不支持资料附件"""
    d = request.get_json(silent=True) or {}
    text_type = str(d.get("type") or "").strip()
    if text_type not in TEXT_RESUME_TYPES:
        return jsonify({"error": "无效的文本项类型"}), 400
    text = str(d.get("text") or "").strip()[:1000]

    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    conn = get_db()
    try:
        if not conn.execute(
            "SELECT 1 FROM students WHERE id = ?", (student_id,)
        ).fetchone():
            return jsonify({"error": "学生档案不存在"}), 404

        existing = conn.execute(
            "SELECT id FROM resume_items WHERE student_id = ? AND type = ? ORDER BY id LIMIT 1",
            (student_id, text_type),
        ).fetchone()

        if text:
            if existing:
                # 文本项固定单条：标题恒为栏目名，仅更新正文
                conn.execute(
                    "UPDATE resume_items SET title = ?, description = ? WHERE id = ?",
                    (TEXT_RESUME_TYPES[text_type], text, existing["id"]),
                )
                item_id = existing["id"]
            else:
                cur = conn.execute(
                    """
                    INSERT INTO resume_items
                        (student_id, type, title, subtitle, date_text, description, created_at)
                    VALUES (?, ?, ?, '', '', ?, ?)
                    """,
                    (student_id, text_type, TEXT_RESUME_TYPES[text_type], text, now),
                )
                item_id = cur.lastrowid
        elif existing:
            # 内容清空则删除该文本项
            conn.execute("DELETE FROM attachments WHERE item_id = ?", (existing["id"],))
            conn.execute("DELETE FROM resume_items WHERE id = ?", (existing["id"],))
            item_id = None
        else:
            item_id = None

        conn.execute(
            "UPDATE students SET updated_at = ? WHERE id = ?", (now, student_id)
        )
        conn.commit()
    finally:
        conn.close()
    return jsonify({"message": "保存成功", "id": item_id})


# ---------- 简历条目（培训/项目/证书/获奖）与资料附件 ----------
def attach_meta(a):
    """为附件附加图标类型与可读大小"""
    a = dict(a)
    ext = os.path.splitext(a.get("original_name") or a.get("filename") or "")[1].lower()
    if ext in {".jpg", ".jpeg", ".png", ".gif", ".webp", ".bmp"}:
        icon = "IMG"
    elif ext in {".zip", ".rar", ".7z", ".tar", ".gz"}:
        icon = "ZIP"
    elif ext == ".pdf":
        icon = "PDF"
    elif ext in {".doc", ".docx"}:
        icon = "DOC"
    elif ext in {".xls", ".xlsx", ".csv"}:
        icon = "XLS"
    elif ext in {".ppt", ".pptx"}:
        icon = "PPT"
    elif ext in {".mp4", ".avi", ".mov", ".mkv", ".wmv"}:
        icon = "VID"
    elif ext in {".mp3", ".wav", ".m4a", ".aac"}:
        icon = "AUD"
    elif ext in {".txt", ".md"}:
        icon = "TXT"
    else:
        icon = "FILE"
    size = a.get("size") or 0
    if size >= 1024 * 1024:
        size_text = f"{size / 1024 / 1024:.1f} MB"
    elif size >= 1024:
        size_text = f"{size / 1024:.0f} KB"
    else:
        size_text = f"{size} B"
    a["icon"] = icon
    a["size_text"] = size_text
    a["is_image"] = icon == "IMG"
    return a


def load_resume_sections(conn, student_id):
    """按类型分组加载学生的简历条目（含附件）"""
    items = [dict(r) for r in conn.execute(
        "SELECT * FROM resume_items WHERE student_id = ? ORDER BY id", (student_id,)
    ).fetchall()]
    atts = {}
    for a in conn.execute(
        "SELECT id, student_id, item_id, filename, original_name, size, uploaded_at FROM attachments ORDER BY item_id, id"
    ).fetchall():
        atts.setdefault(a["item_id"], []).append(attach_meta(a))
    for it in items:
        it["attachments"] = atts.get(it["id"], [])
    sections = {t: [] for t in RESUME_TYPES}
    sections.update({t: [] for t in TEXT_RESUME_TYPES})
    for it in items:
        if it["type"] in sections:
            sections[it["type"]].append(it)
    return sections


@app.get("/api/students/<int:student_id>/resume-items")
def list_resume_items(student_id):
    """获取学生的全部简历条目（按类型分组，含附件）"""
    conn = get_db()
    try:
        if not conn.execute(
            "SELECT 1 FROM students WHERE id = ?", (student_id,)
        ).fetchone():
            return jsonify({"error": "学生档案不存在"}), 404
        sections = load_resume_sections(conn, student_id)
    finally:
        conn.close()
    return jsonify(sections)


@app.post("/api/students/<int:student_id>/resume-items")
@resume_write_required
def create_resume_item(student_id):
    """添加一条简历条目"""
    d = request.get_json(silent=True) or {}
    item_type = str(d.get("type") or "").strip()
    title = str(d.get("title") or "").strip()
    if item_type not in RESUME_TYPES:
        return jsonify({"error": "无效的条目类型"}), 400
    if not title:
        return jsonify({"error": "标题为必填项"}), 400

    conn = get_db()
    try:
        if not conn.execute(
            "SELECT 1 FROM students WHERE id = ?", (student_id,)
        ).fetchone():
            return jsonify({"error": "学生档案不存在"}), 404
        cur = conn.execute(
            """
            INSERT INTO resume_items
                (student_id, type, title, subtitle, date_text, description, created_at)
            VALUES (?, ?, ?, ?, ?, ?, ?)
            """,
            (
                student_id, item_type, title,
                str(d.get("subtitle") or "").strip(),
                str(d.get("date") or "").strip(),
                str(d.get("description") or "").strip(),
                datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            ),
        )
        conn.commit()
        item_id = cur.lastrowid
    finally:
        conn.close()
    return jsonify({"id": item_id, "message": "添加成功"}), 201


@app.put("/api/resume-items/<int:item_id>")
@resume_write_required
def update_resume_item(item_id):
    """修改一条简历条目"""
    d = request.get_json(silent=True) or {}
    title = str(d.get("title") or "").strip()
    if not title:
        return jsonify({"error": "标题为必填项"}), 400

    conn = get_db()
    try:
        cur = conn.execute(
            """
            UPDATE resume_items SET
                title = ?, subtitle = ?, date_text = ?, description = ?
            WHERE id = ?
            """,
            (
                title,
                str(d.get("subtitle") or "").strip(),
                str(d.get("date") or "").strip(),
                str(d.get("description") or "").strip(),
                item_id,
            ),
        )
        conn.commit()
        if cur.rowcount == 0:
            return jsonify({"error": "条目不存在"}), 404
    finally:
        conn.close()
    return jsonify({"message": "修改成功"})


@app.delete("/api/resume-items/<int:item_id>")
@resume_write_required
def delete_resume_item(item_id):
    """删除一条简历条目（附件一并删除）"""
    conn = get_db()
    try:
        if not conn.execute(
            "SELECT 1 FROM resume_items WHERE id = ?", (item_id,)
        ).fetchone():
            return jsonify({"error": "条目不存在"}), 404
        conn.execute("DELETE FROM attachments WHERE item_id = ?", (item_id,))
        conn.execute("DELETE FROM resume_items WHERE id = ?", (item_id,))
        conn.commit()
    finally:
        conn.close()
    return jsonify({"message": "删除成功"})


@app.post("/api/resume-items/<int:item_id>/attachments")
@resume_write_required
def upload_attachments(item_id):
    """为简历条目上传资料附件（图片 / 文案 / 文档 / 压缩包）"""
    files = [f for f in request.files.getlist("files") if f and f.filename]
    if not files:
        return jsonify({"error": "未选择任何文件"}), 400

    conn = get_db()
    try:
        row = conn.execute(
            "SELECT student_id FROM resume_items WHERE id = ?", (item_id,)
        ).fetchone()
        if not row:
            return jsonify({"error": "条目不存在"}), 404
        student_id = row["student_id"]
        existing = conn.execute(
            "SELECT COUNT(*) AS c FROM attachments WHERE item_id = ?", (item_id,)
        ).fetchone()["c"]
        if existing + len(files) > MAX_ATTACHMENTS:
            return jsonify(
                {"error": f"每个条目最多 {MAX_ATTACHMENTS} 个附件（已有 {existing} 个，本次选择 {len(files)} 个）"}
            ), 400

        # 先整体校验大小，全部通过后再统一保存（不限格式）
        blobs = []
        for f in files:
            ext = os.path.splitext(f.filename)[1].lower()
            blob = f.read()
            if len(blob) > MAX_ATT_SIZE:
                return jsonify(
                    {"error": f"文件过大：{f.filename}（单个不超过 50MB）"}
                ), 400
            blobs.append((blob, ext, os.path.basename(f.filename).replace(" ", "_") or "file"))

        now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        for blob, ext, original in blobs:
            stored = f"att_{uuid.uuid4().hex[:16]}{ext}"
            conn.execute(
                "INSERT INTO attachments (student_id, item_id, filename, original_name, size, uploaded_at, data) VALUES (?, ?, ?, ?, ?, ?, ?)",
                (student_id, item_id, stored, original, len(blob), now_str, blob),
            )
        conn.commit()
    finally:
        conn.close()
    return jsonify({"message": f"成功上传 {len(blobs)} 个附件"}), 201


@app.delete("/api/attachments/<int:attachment_id>")
@resume_write_required
def delete_attachment(attachment_id):
    """删除一个附件"""
    conn = get_db()
    try:
        row = conn.execute(
            "SELECT filename FROM attachments WHERE id = ?", (attachment_id,)
        ).fetchone()
        if not row:
            return jsonify({"error": "附件不存在"}), 404
        conn.execute("DELETE FROM attachments WHERE id = ?", (attachment_id,))
        conn.commit()
    finally:
        conn.close()
    return jsonify({"message": "删除成功"})


@app.get("/uploads/files/<path:filename>")
def get_file(filename):
    """获取简历条目附件；带 ?download=1 时以原始文件名下载（只读角色不可下载）"""
    if _viewer_download_denied():
        return jsonify({"error": "只读角色无权下载附件"}), 403
    conn = get_db()
    try:
        row = conn.execute(
            "SELECT data, original_name FROM attachments WHERE filename = ?", (filename,)
        ).fetchone()
    finally:
        conn.close()
    if not row or not row["data"]:
        abort(404)
    download_name = row["original_name"] or filename
    return send_file(
        io.BytesIO(row["data"]),
        download_name=download_name,
        as_attachment=bool(request.args.get("download")),
    )


# ---------- 考勤管理 ----------
ATT_STATUS = {"present": "出勤", "late": "迟到", "leave": "请假", "absent": "缺勤"}


@app.route("/attendance")
@login_required
def attendance_page():
    """考勤管理页"""
    return render_template("attendance.html", user=current_user())


@app.get("/api/attendance")
def get_attendance():
    """获取某日期的考勤记录"""
    date = request.args.get("date", "").strip()
    if not re.match(r"^\d{4}-\d{2}-\d{2}$", date):
        return jsonify({"error": "日期格式不正确"}), 400
    conn = get_db()
    try:
        rows = conn.execute(
            """
            SELECT id, student_id, date, status, remark,
                   course_name, lesson, teacher, start_time, end_time
            FROM attendance WHERE date = ? ORDER BY student_id
            """,
            (date,),
        ).fetchall()
    finally:
        conn.close()
    return jsonify([dict(r) for r in rows])


@app.post("/api/attendance")
@permission_required("attendance.save", "当前角色无权保存考勤记录")
def save_attendance():
    """批量保存某日期考勤（同一学生同一天仅一条，重复保存自动覆盖）"""
    d = request.get_json(silent=True) or {}
    date = str(d.get("date") or "").strip()
    records = d.get("records") or []
    if not re.match(r"^\d{4}-\d{2}-\d{2}$", date):
        return jsonify({"error": "日期格式不正确"}), 400
    if not records:
        return jsonify({"error": "没有需要保存的考勤记录"}), 400
    if len(records) > 1000:
        return jsonify({"error": "单次保存记录过多"}), 400

    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    conn = get_db()
    try:
        saved = 0
        for r in records:
            if not isinstance(r, dict):
                continue
            student_id = r.get("student_id")
            status = str(r.get("status") or "").strip()
            remark = str(r.get("remark") or "").strip()[:100]
            course_name = str(r.get("course_name") or "").strip()[:100]
            lesson = str(r.get("lesson") or "").strip()[:100]
            teacher = str(r.get("teacher") or "").strip()[:50]
            start_time = str(r.get("start_time") or "").strip()[:5]
            end_time = str(r.get("end_time") or "").strip()[:5]
            if not re.match(r"^\d{2}:\d{2}$", start_time):
                start_time = ""
            if not re.match(r"^\d{2}:\d{2}$", end_time):
                end_time = ""
            source = str(r.get("source") or "").strip()
            if source not in ("", "makeup"):
                source = ""
            if not isinstance(student_id, int) or status not in ATT_STATUS:
                continue
            if not conn.execute(
                "SELECT 1 FROM students WHERE id = ?", (student_id,)
            ).fetchone():
                continue
            conn.execute(
                """
                INSERT INTO attendance
                    (student_id, date, status, remark, course_name, lesson, teacher,
                     start_time, end_time, source, created_at, updated_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT(student_id, date) DO UPDATE SET
                    status = excluded.status,
                    remark = excluded.remark,
                    course_name = excluded.course_name,
                    lesson = excluded.lesson,
                    teacher = excluded.teacher,
                    start_time = excluded.start_time,
                    end_time = excluded.end_time,
                    source = excluded.source,
                    updated_at = excluded.updated_at
                """,
                (student_id, date, status, remark, course_name, lesson, teacher,
                 start_time, end_time, source, now_str, now_str),
            )
            saved += 1
        conn.commit()
    finally:
        conn.close()
    return jsonify({"message": f"已保存 {saved} 条考勤记录"})


def _history_filter(start, end, status, keyword, source=""):
    """构造考勤记录查询条件（记录查询与 Excel 导出共用）；source=makeup 时只查补考勤记录"""
    where, params = [], []
    if re.match(r"^\d{4}-\d{2}-\d{2}$", start):
        where.append("a.date >= ?")
        params.append(start)
    if re.match(r"^\d{4}-\d{2}-\d{2}$", end):
        where.append("a.date <= ?")
        params.append(end)
    if source:
        where.append("a.source = ?")
        params.append(source)
    # 状态多选：逗号分隔（如 late,absent），逐一校验；未选或全无效 = 全部状态
    statuses = [s for s in (x.strip() for x in status.split(",")) if s in ATT_STATUS]
    if statuses:
        where.append("a.status IN (" + ",".join("?" * len(statuses)) + ")")
        params.extend(statuses)
    if keyword:   # 关键字同时匹配学号与学生姓名
        where.append("(s.name LIKE ? OR s.student_no LIKE ?)")
        params.append(f"%{keyword}%")
        params.append(f"%{keyword}%")
    cond = ("WHERE " + " AND ".join(where)) if where else ""
    return cond, params


def _history_rows_stats(cond, params):
    """按查询条件读取考勤明细与各状态统计"""
    conn = get_db()
    try:
        rows = conn.execute(
            f"""
            SELECT a.id, a.student_id, a.date, a.status, a.remark,
                   a.course_name, a.lesson, a.teacher, a.start_time, a.end_time,
                   s.student_no, s.name
            FROM attendance a JOIN students s ON s.id = a.student_id
            {cond}
            ORDER BY a.date DESC, a.student_id
            LIMIT 1000
            """,
            params,
        ).fetchall()
        st = conn.execute(
            f"""
            SELECT a.status, COUNT(*) AS c
            FROM attendance a JOIN students s ON s.id = a.student_id
            {cond}
            GROUP BY a.status
            """,
            params,
        ).fetchall()
    finally:
        conn.close()
    stats = {k: 0 for k in ATT_STATUS}
    for r in st:
        if r["status"] in stats:
            stats[r["status"]] = r["c"]
    stats["total"] = sum(stats.values())
    return rows, stats


@app.get("/api/attendance/history")
def attendance_history():
    """考勤记录查询（日期范围 + 状态 + 学生关键字），含统计"""
    cond, params = _history_filter(
        request.args.get("start", "").strip(),
        request.args.get("end", "").strip(),
        request.args.get("status", "").strip(),
        request.args.get("keyword", "").strip(),
        request.args.get("source", "").strip(),
    )
    rows, stats = _history_rows_stats(cond, params)
    return jsonify({"records": [dict(r) for r in rows], "stats": stats})


@app.get("/api/attendance/export")
@permission_required("attendance.export", "当前角色无权导出考勤 Excel")
def export_attendance():
    """考勤记录导出 Excel：明细后附统计，末行写出总课时（仅计出勤、迟到）"""
    start = request.args.get("start", "").strip()
    end = request.args.get("end", "").strip()
    status = request.args.get("status", "").strip()
    keyword = request.args.get("keyword", "").strip()
    cond, params = _history_filter(start, end, status, keyword)
    rows, stats = _history_rows_stats(cond, params)

    wb = Workbook()
    ws = wb.active
    ws.title = "考勤记录"

    thin = Side(style="thin", color="D9D9D9")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    head_fill = PatternFill("solid", fgColor="4472C4")
    sum_fill = PatternFill("solid", fgColor="FFF2CC")
    head_font = Font(name="等线", bold=True, color="FFFFFF")
    title_font = Font(name="等线", size=14, bold=True)
    sub_font = Font(name="等线", size=9, color="808080")
    cell_font = Font(name="等线")
    center = Alignment(horizontal="center", vertical="center")
    left = Alignment(horizontal="left", vertical="center", wrap_text=True)

    # 标题与筛选说明
    ws.merge_cells("A1:I1")
    ws["A1"] = "考勤记录明细"
    ws["A1"].font = title_font
    ws["A1"].alignment = center
    ws.merge_cells("A2:I2")
    # 附加学员与第一位家长信息：学员取导出结果第一位；家长取学生档案中
    # 该学员的第一条监护人记录（guardians 表同一学员多条时按录入顺序取最前一条）
    stu_name, g_name, g_phone = "无", "无", "无"
    if rows:
        first = rows[0]
        stu_name = first["name"] or "无"
        conn = get_db()
        try:
            g = conn.execute(
                "SELECT name, phone FROM guardians WHERE student_id = ? ORDER BY id LIMIT 1",
                (first["student_id"],),
            ).fetchone()
        finally:
            conn.close()
        if g:
            g_name = g["name"] or "无"
            g_phone = g["phone"] or "无"
    ws["A2"] = (
        f"日期范围：{start or '不限'} 至 {end or '不限'}　"
        f"导出时间：{datetime.now().strftime('%Y-%m-%d %H:%M')}　"
        f"学员姓名：{stu_name}　家长姓名：{g_name}　家长手机号：{g_phone}"
    )
    ws["A2"].font = sub_font
    ws["A2"].alignment = left
    ws.row_dimensions[1].height = 24
    ws.row_dimensions[2].height = 28

    # 表头
    head_row = 4
    headers = ["日期", "学号", "姓名", "课程", "本节课", "上课时间", "老师", "状态", "备注"]
    widths = [12, 10, 12, 20, 22, 13, 12, 8, 26]
    for i, (h, w) in enumerate(zip(headers, widths), start=1):
        c = ws.cell(row=head_row, column=i, value=h)
        c.font = head_font
        c.fill = head_fill
        c.alignment = center
        c.border = border
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A5"

    # 数据行
    row = head_row + 1
    for r in rows:
        vals = [
            r["date"], r["student_no"] or "", r["name"] or "",
            r["course_name"] or "", r["lesson"] or "",
            ((r["start_time"] or "") + (("~" + r["end_time"]) if r["end_time"] else "")),
            r["teacher"] or "",
            ATT_STATUS.get(r["status"], r["status"]), r["remark"] or "",
        ]
        for col, v in enumerate(vals, start=1):
            c = ws.cell(row=row, column=col, value=v)
            c.font = cell_font
            c.border = border
            c.alignment = center if col in (1, 2, 8) else left
        row += 1

    # 统计汇总 + 总课时（总课时 = 出勤 + 迟到，请假和缺勤不计入）
    row += 1
    ws.cell(row=row, column=1, value="统计汇总").font = Font(name="等线", bold=True)
    row += 1
    for label, v in (
        ("出勤人次", stats["present"]),
        ("迟到人次", stats["late"]),
        ("请假人次", stats["leave"]),
        ("缺勤人次", stats["absent"]),
        ("记录总条数", stats["total"]),
    ):
        ws.cell(row=row, column=1, value=label).font = cell_font
        ws.cell(row=row, column=2, value=v).font = cell_font
        row += 1

    total_hours = stats["present"] + stats["late"]
    ws.cell(row=row, column=1, value="总课时").font = Font(name="等线", bold=True)
    hc = ws.cell(row=row, column=2, value=total_hours)
    hc.font = Font(name="等线", bold=True, color="C00000")
    for col in (1, 2):
        ws.cell(row=row, column=col).fill = sum_fill
    note = ws.cell(row=row, column=3,
                   value="总课时 = 出勤 + 迟到；请假和缺勤不计入总课时")
    note.font = sub_font
    note.alignment = left
    ws.merge_cells(start_row=row, start_column=3, end_row=row, end_column=8)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return send_file(
        buf,
        as_attachment=True,
        download_name=datetime.now().strftime("考勤记录_%Y%m%d_%H%M%S.xlsx"),
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.delete("/api/attendance/<int:record_id>")
@permission_required("attendance.delete", "教师不可删除或补录考勤数据")
def delete_attendance(record_id):
    """删除一条考勤记录"""
    conn = get_db()
    try:
        cur = conn.execute("DELETE FROM attendance WHERE id = ?", (record_id,))
        conn.commit()
        if cur.rowcount == 0:
            return jsonify({"error": "记录不存在"}), 404
    finally:
        conn.close()
    return jsonify({"message": "删除成功"})


@app.post("/api/attendance/batch")
@permission_required("attendance.batch", "教师无权批量修改考勤记录")
def batch_update_attendance():
    """批量修改考勤记录（记录查询页「修改查询信息状态」）

    按记录 id 批量更新状态与/或备注：
    - status 传空字符串表示不修改状态；传值时必须是四种状态之一
    - remark 传 null 表示不修改备注；传字符串则覆盖（支持传空字符串清空）
    """
    d = request.get_json(silent=True) or {}
    ids = []
    for i in d.get("ids") or []:
        try:
            n = int(i)
        except (TypeError, ValueError):
            continue
        if n > 0 and n not in ids:
            ids.append(n)
    if not ids:
        return jsonify({"error": "请先选择要修改的考勤记录"}), 400
    if len(ids) > 1000:
        return jsonify({"error": "单次最多修改 1000 条记录"}), 400

    status = str(d.get("status") or "").strip()
    if status and status not in ATT_STATUS:
        return jsonify({"error": "考勤状态不正确"}), 400
    remark = d.get("remark")
    remark_given = remark is not None          # 未传备注 = 保持原备注不动
    remark = str(remark).strip()[:100] if remark_given else ""
    if not status and not remark_given:
        return jsonify({"error": "请填写要修改的状态或备注"}), 400

    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    sets, params = [], []
    if status:
        sets.append("status = ?")
        params.append(status)
    if remark_given:
        sets.append("remark = ?")
        params.append(remark)
    sets.append("updated_at = ?")
    params.append(now_str)
    params.extend(ids)
    sql = (
        "UPDATE attendance SET " + ", ".join(sets)
        + " WHERE id IN (" + ",".join("?" * len(ids)) + ")"
    )
    conn = get_db()
    try:
        cur = conn.execute(sql, params)
        conn.commit()
        affected = cur.rowcount
    finally:
        conn.close()
    return jsonify({"message": f"已修改 {affected} 条考勤记录", "affected": affected})


# ---------- 课时管理 ----------
def _lesson_stats_rows(keyword="", warn_only=False, class_id=None, teacher_id=None):
    """课时数据公共查询：学生维度汇总（已购/已退/已消耗/剩余/消耗率/预警）

    支持按班级（课程报读）与授课教师筛选：教师名下所有班级的学员。
    """
    where, params = [], []
    if keyword:
        where.append("(s.name LIKE ? OR s.student_no LIKE ?)")
        params += [f"%{keyword}%", f"%{keyword}%"]
    if warn_only:
        # 预警 = 剩余课时 <= 0（课时已用尽或超上）
        where.append("(COALESCE(p_t.t, 0) - COALESCE(r_t.t, 0) - COALESCE(a_t.c, 0)) <= 0")
    if class_id:
        # 班级筛选：该班级（课程）报读的学员
        where.append("s.id IN (SELECT e.student_id FROM enrollments e WHERE e.course_id = ?)")
        params.append(class_id)
    if teacher_id:
        # 教师筛选：其授课的 所有班级 的学员（兼容 teacher_id 外键与 teacher 姓名两种关联）
        where.append(
            """s.id IN (SELECT e.student_id FROM enrollments e
                       JOIN courses c ON c.id = e.course_id
                       WHERE c.teacher_id = ?
                          OR c.teacher = (SELECT t.name FROM teachers t WHERE t.id = ?))"""
        )
        params += [teacher_id, teacher_id]
    cond = ("WHERE " + " AND ".join(where)) if where else ""
    sql = f"""
        SELECT s.id, s.student_no, s.name AS student_name, s.gender, s.grade, s.course,
               COALESCE(p_t.t, 0) AS total_lessons,
               COALESCE(r_t.t, 0) AS refunded_lessons,
               COALESCE(a_t.c, 0) AS used_lessons,
               COALESCE(p_t.t, 0) - COALESCE(r_t.t, 0) - COALESCE(a_t.c, 0) AS remaining_lessons,
               COALESCE(p_t.p, 0) AS paid_amount,
               COALESCE(r_t.a, 0) AS refunded_amount,
               (SELECT GROUP_CONCAT(c.name, '、') FROM enrollments e
                JOIN courses c ON c.id = e.course_id
                WHERE e.student_id = s.id) AS class_names,
               (SELECT GROUP_CONCAT(t_name, '、') FROM (
                    SELECT DISTINCT c.teacher AS t_name FROM enrollments e
                    JOIN courses c ON c.id = e.course_id
                    WHERE e.student_id = s.id AND c.teacher IS NOT NULL AND c.teacher != ''
                )) AS class_teachers
        FROM students s
        LEFT JOIN (SELECT student_id, SUM(lesson_count) AS t, SUM(amount) AS p
                   FROM payments GROUP BY student_id) p_t ON p_t.student_id = s.id
        LEFT JOIN (SELECT student_id, SUM(lesson_count) AS t, SUM(amount) AS a
                   FROM refunds GROUP BY student_id) r_t ON r_t.student_id = s.id
        LEFT JOIN (SELECT student_id, COUNT(*) AS c
                   FROM attendance WHERE status IN ('present', 'late')
                   GROUP BY student_id) a_t ON a_t.student_id = s.id
        {cond}
        ORDER BY remaining_lessons ASC, s.id ASC
        LIMIT 5000
    """
    conn = get_db()
    try:
        rows = [dict(r) for r in conn.execute(sql, params).fetchall()]
    finally:
        conn.close()
    for r in rows:
        valid = r["total_lessons"] - r["refunded_lessons"]
        r["use_rate"] = round(r["used_lessons"] * 100 / valid, 1) if valid > 0 else 0
        # 预警 = 剩余课时 <= 0（含未购课但已上课的情况）
        r["warn"] = r["remaining_lessons"] <= 0
    return rows


@app.route("/lessons")
@login_required
def lessons_page():
    """课时管理页"""
    return render_template("lessons.html", user=current_user())


@app.get("/api/lessons/filters")
def lessons_filters():
    """课时管理筛选选项：班级（课程）与授课教师下拉数据"""
    conn = get_db()
    try:
        classes = [
            dict(r) for r in conn.execute(
                """
                SELECT c.id, c.name, c.teacher, c.status,
                       (SELECT COUNT(*) FROM enrollments e WHERE e.course_id = c.id) AS enrolled
                FROM courses c
                ORDER BY c.status = '开课中' DESC, c.id DESC
                """
            ).fetchall()
        ]
        teachers = [
            dict(r) for r in conn.execute(
                """
                SELECT DISTINCT t.id, t.name, t.subject, t.status
                FROM teachers t
                JOIN courses c ON (c.teacher_id = t.id OR c.teacher = t.name)
                ORDER BY t.status = '在职' DESC, t.id ASC
                """
            ).fetchall()
        ]
    finally:
        conn.close()
    return jsonify({"classes": classes, "teachers": teachers})


@app.get("/api/lessons")
def list_lessons():
    """学员课时信息列表，支持关键字/班级/授课教师查询与课时预警筛选"""
    keyword = request.args.get("keyword", "").strip()
    warn_only = request.args.get("warn") == "1"
    class_id = request.args.get("class_id", type=int) or None
    teacher_id = request.args.get("teacher_id", type=int) or None
    rows = _lesson_stats_rows(keyword, warn_only, class_id=class_id, teacher_id=teacher_id)
    total_bought = sum(r["total_lessons"] for r in rows)
    total_refunded = sum(r["refunded_lessons"] for r in rows)
    total_used = sum(r["used_lessons"] for r in rows)
    total_remaining = sum(r["remaining_lessons"] for r in rows)
    return jsonify({
        "rows": rows,
        "stats": {
            "students": len(rows),
            "total_bought": total_bought,
            "total_refunded": total_refunded,
            "total_used": total_used,
            "total_remaining": total_remaining,
            "warn_count": sum(1 for r in rows if r["warn"]),
        },
    })


@app.get("/api/lessons/export")
@permission_required("lesson.export", "当前角色无权导出课时 Excel")
def export_lessons():
    """学员课时信息导出 Excel"""
    keyword = request.args.get("keyword", "").strip()
    warn_only = request.args.get("warn") == "1"
    class_id = request.args.get("class_id", type=int) or None
    teacher_id = request.args.get("teacher_id", type=int) or None
    rows = _lesson_stats_rows(keyword, warn_only, class_id=class_id, teacher_id=teacher_id)

    wb = Workbook()
    ws = wb.active
    ws.title = "学员课时"

    thin = Side(style="thin", color="D9D9D9")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    head_fill = PatternFill("solid", fgColor="4472C4")
    sum_fill = PatternFill("solid", fgColor="FFF2CC")
    head_font = Font(name="等线", bold=True, color="FFFFFF")
    title_font = Font(name="等线", size=14, bold=True)
    sub_font = Font(name="等线", size=9, color="808080")
    cell_font = Font(name="等线")
    center = Alignment(horizontal="center", vertical="center")

    # 标题与说明
    ws.merge_cells("A1:M1")
    ws["A1"] = "学员课时信息表"
    ws["A1"].font = title_font
    ws["A1"].alignment = center
    ws.row_dimensions[1].height =  30
    ws.merge_cells("A2:M2")
    ws["A2"] = (f"筛选：{('关键字「' + keyword + '」') if keyword else '全部学员'}"
                f"{' · 仅课时预警' if warn_only else ''}　"
                f"导出时间：{datetime.now().strftime('%Y-%m-%d %H:%M')}")
    ws["A2"].font = sub_font
    ws["A2"].alignment = center
    ws.row_dimensions[2].height = 20

    # 表头（剩余课时紧跟姓名，核心列前置；含班级与授课教师）
    headers = ["学号", "姓名", "剩余课时", "课时使用率", "课时状态", "已购课时",
               "已退课时", "已消耗课时", "所属班级", "授课教师", "性别", "年级", "报读科目"]
    widths = [10, 10, 10, 12, 10, 10, 10, 12, 18, 14, 6, 10, 16]
    head_row = 4
    for i, (h, w) in enumerate(zip(headers, widths), start=1):
        c = ws.cell(row=head_row, column=i, value=h)
        c.font = head_font
        c.fill = head_fill
        c.alignment = center
        c.border = border
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A5"

    # 数据行
    row = head_row + 1
    for r in rows:
        vals = [
            r["student_no"] or "", r["student_name"] or "",
            r["remaining_lessons"], f"{r['use_rate']}%",
            "课时预警" if r["warn"] else ("正常" if r["total_lessons"] > 0 else "未购课"),
            r["total_lessons"], r["refunded_lessons"], r["used_lessons"],
            r.get("class_names") or "", r.get("class_teachers") or "",
            r["gender"] or "", r["grade"] or "", r["course"] or "",
        ]
        for col, v in enumerate(vals, start=1):
            c = ws.cell(row=row, column=col, value=v)
            c.font = cell_font
            c.border = border
            c.alignment = center
            if col == 3:
                c.font = Font(name="等线", bold=True,
                              color="C00000" if r["remaining_lessons"] <= 0 else "2F5597")
        row += 1

    # 汇总
    row += 1
    ws.cell(row=row, column=1, value="统计汇总").font = Font(name="等线", bold=True)
    row += 1
    total_bought = sum(r["total_lessons"] for r in rows)
    total_refunded = sum(r["refunded_lessons"] for r in rows)
    total_used = sum(r["used_lessons"] for r in rows)
    total_remaining = sum(r["remaining_lessons"] for r in rows)
    for label, v in (("学员人数", len(rows)), ("已购课时合计", total_bought),
                     ("已退课时合计", total_refunded), ("已消耗课时合计", total_used),
                     ("剩余课时合计", total_remaining)):
        ws.cell(row=row, column=1, value=label).font = cell_font
        ws.cell(row=row, column=2, value=v).font = cell_font
        for col in (1, 2):
            ws.cell(row=row, column=col).fill = sum_fill
        row += 1
    note = ws.cell(row=row, column=1,
                   value="剩余课时 = 已购课时 − 已退课时 − 已消耗课时；已消耗 = 出勤 + 迟到（请假/缺勤不扣课时）")
    note.font = sub_font
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=13)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return send_file(
        buf,
        as_attachment=True,
        download_name=datetime.now().strftime("学员课时表_%Y%m%d_%H%M%S.xlsx"),
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


# ---------- 课程管理 ----------
COURSE_STATUS = ("开课中", "已结课")


def _to_int(v):
    """转为非负整数，失败返回 None"""
    try:
        n = int(str(v).strip())
        return n if n >= 0 else None
    except (ValueError, TypeError):
        return None


def _to_float(v):
    """转为非负数值，失败返回 None"""
    try:
        n = float(str(v).strip())
        return n if n >= 0 else None
    except (ValueError, TypeError):
        return None


@app.route("/courses")
@login_required
def courses_page():
    """课程管理页"""
    return render_template("courses.html", user=current_user())


@app.get("/api/courses")
def list_courses():
    """课程列表，支持关键字（名称/教师/类别）与状态筛选，含报读人数"""
    keyword = request.args.get("keyword", "").strip()
    status = request.args.get("status", "").strip()
    where, params = [], []
    if keyword:
        like = f"%{keyword}%"
        where.append("(c.name LIKE ? OR c.teacher LIKE ? OR c.category LIKE ?)")
        params += [like, like, like]
    if status in COURSE_STATUS:
        where.append("c.status = ?")
        params.append(status)
    cond = ("WHERE " + " AND ".join(where)) if where else ""

    conn = get_db()
    try:
        rows = conn.execute(
            f"""
            SELECT c.*,
                   (SELECT COUNT(*) FROM enrollments e WHERE e.course_id = c.id) AS enrolled,
                   (SELECT t.teacher_no FROM teachers t WHERE t.name = c.teacher ORDER BY t.id LIMIT 1) AS teacher_no
            FROM courses c
            {cond}
            ORDER BY c.id DESC
            """,
            params,
        ).fetchall()

        # 附带课程学员（id / 学号 / 姓名）
        stu_map = {}
        for e in conn.execute(
            """
            SELECT e.course_id, s.id AS sid, s.student_no, s.name, s.gender,
                   (SELECT p.filename FROM photos p
                    WHERE p.student_id = s.id ORDER BY p.sort, p.id LIMIT 1) AS photo
            FROM enrollments e JOIN students s ON s.id = e.student_id
            ORDER BY e.course_id, e.id
            """
        ).fetchall():
            stu_map.setdefault(e["course_id"], []).append(
                {
                    "id": e["sid"], "student_no": e["student_no"], "name": e["name"],
                    "gender": e["gender"], "photo": e["photo"],
                }
            )
    finally:
        conn.close()

    courses = []
    for r in rows:
        c = dict(r)
        c["students"] = stu_map.get(c["id"], [])
        courses.append(c)
    return jsonify(courses)


@app.post("/api/courses")
@permission_required("course.write", "只读角色无权新建课程")
def create_course():
    """新建课程"""
    d = request.get_json(silent=True) or {}
    name = str(d.get("name") or "").strip()
    if not name:
        return jsonify({"error": "课程名称为必填项"}), 400
    status = str(d.get("status") or "").strip()
    if status not in COURSE_STATUS:
        status = "开课中"

    start_date = str(d.get("start_date") or "").strip()
    end_date = str(d.get("end_date") or "").strip()
    if not start_date:
        start_date = datetime.now().strftime("%Y-%m-%d")   # 开课日期默认为创建当天
    for label, val in (("开课", start_date), ("结课", end_date)):
        if val and not re.fullmatch(r"\d{4}-\d{2}-\d{2}", val):
            return jsonify({"error": f"{label}日期格式不正确（应为 YYYY-MM-DD）"}), 400
    if end_date and end_date < start_date:
        return jsonify({"error": "结课日期不能早于开课日期"}), 400

    conn = get_db()
    try:
        cur = conn.execute(
            """
            INSERT INTO courses
                (name, category, teacher, teacher_id, schedule, hours, price, capacity,
                 start_date, end_date, description, status, created_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                name,
                str(d.get("category") or "").strip(),
                str(d.get("teacher") or "").strip(),
                d.get("teacher_id"),
                str(d.get("schedule") or "").strip(),
                _to_int(d.get("hours")),
                _to_float(d.get("price")),
                _to_int(d.get("capacity")),
                start_date,
                end_date,
                str(d.get("description") or "").strip(),
                status,
                datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            ),
        )
        conn.commit()
        course_id = cur.lastrowid
    finally:
        conn.close()
    return jsonify({"id": course_id, "message": "创建成功"}), 201


@app.put("/api/courses/<int:course_id>")
@permission_required("course.write", "只读角色无权修改课程")
def update_course(course_id):
    """修改课程"""
    d = request.get_json(silent=True) or {}
    name = str(d.get("name") or "").strip()
    if not name:
        return jsonify({"error": "课程名称为必填项"}), 400
    status = str(d.get("status") or "").strip()
    if status not in COURSE_STATUS:
        status = "开课中"

    start_date = str(d.get("start_date") or "").strip()
    end_date = str(d.get("end_date") or "").strip()
    for label, val in (("开课", start_date), ("结课", end_date)):
        if val and not re.fullmatch(r"\d{4}-\d{2}-\d{2}", val):
            return jsonify({"error": f"{label}日期格式不正确（应为 YYYY-MM-DD）"}), 400
    if end_date and start_date and end_date < start_date:
        return jsonify({"error": "结课日期不能早于开课日期"}), 400

    conn = get_db()
    try:
        old = conn.execute(
            "SELECT start_date, end_date, created_at FROM courses WHERE id = ?",
            (course_id,),
        ).fetchone()
        if old is None:
            return jsonify({"error": "课程不存在"}), 404
        # 开课日期留空时沿用原值（原值也缺失则取创建当天，与新建规则一致）
        if not start_date:
            start_date = old["start_date"] or (old["created_at"] or "")[:10]
            start_date = start_date or datetime.now().strftime("%Y-%m-%d")
        conn.execute(
            """
            UPDATE courses SET
                name = ?, category = ?, teacher = ?, teacher_id = ?, schedule = ?,
                hours = ?, price = ?, capacity = ?, start_date = ?, end_date = ?,
                description = ?, status = ?
            WHERE id = ?
            """,
            (
                name,
                str(d.get("category") or "").strip(),
                str(d.get("teacher") or "").strip(),
                d.get("teacher_id"),
                str(d.get("schedule") or "").strip(),
                _to_int(d.get("hours")),
                _to_float(d.get("price")),
                _to_int(d.get("capacity")),
                start_date,
                end_date,
                str(d.get("description") or "").strip(),
                status,
                course_id,
            ),
        )
        conn.commit()
    finally:
        conn.close()
    return jsonify({"message": "修改成功"})


@app.delete("/api/courses/<int:course_id>")
@permission_required("course.write", "只读角色无权删除课程")
def delete_course(course_id):
    """删除课程"""
    conn = get_db()
    try:
        cur = conn.execute("DELETE FROM courses WHERE id = ?", (course_id,))
        conn.execute("DELETE FROM enrollments WHERE course_id = ?", (course_id,))
        conn.commit()
        if cur.rowcount == 0:
            return jsonify({"error": "课程不存在"}), 404
    finally:
        conn.close()
    return jsonify({"message": "删除成功"})


@app.post("/api/courses/<int:course_id>/students")
@permission_required("course.write", "只读角色无权修改课程学员名单")
def sync_course_students(course_id):
    """同步课程学员名单（全量覆盖，学员来自学生档案）"""
    d = request.get_json(silent=True) or {}
    ids = d.get("student_ids") or []
    if not isinstance(ids, list) or len(ids) > 500:
        return jsonify({"error": "学员列表不合法"}), 400

    conn = get_db()
    try:
        if not conn.execute(
            "SELECT 1 FROM courses WHERE id = ?", (course_id,)
        ).fetchone():
            return jsonify({"error": "课程不存在"}), 404
        valid_ids = []
        for sid in ids:
            if isinstance(sid, int) and sid not in valid_ids and conn.execute(
                "SELECT 1 FROM students WHERE id = ?", (sid,)
            ).fetchone():
                valid_ids.append(sid)
        conn.execute("DELETE FROM enrollments WHERE course_id = ?", (course_id,))
        now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        for sid in valid_ids:
            conn.execute(
                "INSERT INTO enrollments (course_id, student_id, created_at) VALUES (?, ?, ?)",
                (course_id, sid, now_str),
            )
        conn.commit()
    finally:
        conn.close()
    return jsonify({"message": f"已保存 {len(valid_ids)} 名学员"})


# ---------- 教师管理 ----------
TEACHER_STATUS = ("在职", "离职")


@app.route("/teachers")
@login_required
def teachers_page():
    """教师管理页"""
    return render_template("teachers.html", user=current_user())


@app.get("/api/teachers")
def list_teachers():
    """教师列表，支持关键字（姓名/科目/电话）与状态筛选，含授课门数"""
    keyword = request.args.get("keyword", "").strip()
    status = request.args.get("status", "").strip()
    where, params = [], []
    if keyword:
        like = f"%{keyword}%"
        where.append("(t.teacher_no LIKE ? OR t.name LIKE ? OR t.subject LIKE ? OR t.phone LIKE ?)")
        params += [like, like, like, like]
    if status in TEACHER_STATUS:
        where.append("t.status = ?")
        params.append(status)
    cond = ("WHERE " + " AND ".join(where)) if where else ""

    conn = get_db()
    try:
        rows = conn.execute(
            f"""
            SELECT t.*,
                   (SELECT COUNT(*) FROM courses c WHERE c.teacher = t.name) AS course_count
            FROM teachers t
            {cond}
            ORDER BY t.id DESC
            """,
            params,
        ).fetchall()
    finally:
        conn.close()
    return jsonify([dict(r) for r in rows])


@app.post("/api/teachers")
@permission_required("teacher.create", "当前角色无权新建教师档案")
def create_teacher():
    """新建教师"""
    d = request.get_json(silent=True) or {}
    name = str(d.get("name") or "").strip()
    if not name:
        return jsonify({"error": "教师姓名为必填项"}), 400
    gender = str(d.get("gender") or "").strip()
    if gender not in ("男", "女"):
        gender = "男"
    status = str(d.get("status") or "").strip()
    if status not in TEACHER_STATUS:
        status = "在职"

    conn = get_db()
    try:
        teacher_no = next_teacher_no(conn)
        cur = conn.execute(
            """
            INSERT INTO teachers
                (teacher_no, name, gender, phone, subject, title, hire_date, status, notes, created_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                teacher_no, name, gender,
                str(d.get("phone") or "").strip(),
                str(d.get("subject") or "").strip(),
                str(d.get("title") or "").strip(),
                str(d.get("hire_date") or "").strip(),
                status,
                str(d.get("notes") or "").strip(),
                datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            ),
        )
        conn.commit()
        teacher_id = cur.lastrowid
    finally:
        conn.close()
    return jsonify({"id": teacher_id, "teacher_no": teacher_no, "message": "创建成功"}), 201


@app.put("/api/teachers/<int:teacher_id>")
@permission_required("teacher.write", "教务不可修改教师档案，修改需由管理员或校长操作")
def update_teacher(teacher_id):
    """修改教师信息"""
    d = request.get_json(silent=True) or {}
    name = str(d.get("name") or "").strip()
    if not name:
        return jsonify({"error": "教师姓名为必填项"}), 400
    gender = str(d.get("gender") or "").strip()
    if gender not in ("男", "女"):
        gender = "男"
    status = str(d.get("status") or "").strip()
    if status not in TEACHER_STATUS:
        status = "在职"

    conn = get_db()
    try:
        cur = conn.execute(
            """
            UPDATE teachers SET
                name = ?, gender = ?, phone = ?, subject = ?,
                title = ?, hire_date = ?, status = ?, notes = ?
            WHERE id = ?
            """,
            (
                name, gender,
                str(d.get("phone") or "").strip(),
                str(d.get("subject") or "").strip(),
                str(d.get("title") or "").strip(),
                str(d.get("hire_date") or "").strip(),
                status,
                str(d.get("notes") or "").strip(),
                teacher_id,
            ),
        )
        conn.commit()
        if cur.rowcount == 0:
            return jsonify({"error": "教师不存在"}), 404
    finally:
        conn.close()
    return jsonify({"message": "修改成功"})


@app.delete("/api/teachers/<int:teacher_id>")
@permission_required("teacher.write", "教务不可删除教师档案，删除需由管理员或校长操作")
def delete_teacher(teacher_id):
    """删除教师"""
    conn = get_db()
    try:
        cur = conn.execute("DELETE FROM teachers WHERE id = ?", (teacher_id,))
        conn.commit()
        if cur.rowcount == 0:
            return jsonify({"error": "教师不存在"}), 404
    finally:
        conn.close()
    return jsonify({"message": "删除成功"})


# ---------- 缴费管理 ----------
PAY_METHODS = ("现金", "微信", "支付宝", "银行转账", "其他")


@app.route("/payments")
@login_required
def payments_page():
    """缴费管理页"""
    return render_template("payments.html", user=current_user())


@app.get("/api/payments")
def list_payments():
    """缴费记录，支持日期范围与学生姓名筛选"""
    start = request.args.get("start", "").strip()
    end = request.args.get("end", "").strip()
    keyword = request.args.get("keyword", "").strip()
    where, params = [], []
    if re.match(r"^\d{4}-\d{2}-\d{2}$", start):
        where.append("p.pay_date >= ?")
        params.append(start)
    if re.match(r"^\d{4}-\d{2}-\d{2}$", end):
        where.append("p.pay_date <= ?")
        params.append(end)
    if keyword:
        where.append("(s.name LIKE ? OR s.student_no LIKE ?)")
        params.append(f"%{keyword}%")
        params.append(f"%{keyword}%")
    cond = ("WHERE " + " AND ".join(where)) if where else ""

    conn = get_db()
    try:
        rows = conn.execute(
            f"""
            SELECT p.id, p.student_id, p.course, p.amount, p.pay_date, p.method, p.remark, p.status,
                   p.unit_price, p.lesson_count, p.discount_percent, p.discount_amount,
                   s.student_no, s.name AS student_name
            FROM payments p JOIN students s ON s.id = p.student_id
            {cond}
            ORDER BY p.pay_date DESC, p.id DESC
            LIMIT 1000
            """,
            params,
        ).fetchall()
        result = _attach_photos(conn, rows, "payment_photos", "payment_id")
    finally:
        conn.close()
    return jsonify(result)


@app.get("/api/payments/query")
def query_payments():
    """缴费查询，支持日期范围、学生姓名、科目、缴费方式筛选"""
    start = request.args.get("start", "").strip()
    end = request.args.get("end", "").strip()
    keyword = request.args.get("keyword", "").strip()
    course = request.args.get("course", "").strip()
    method = request.args.get("method", "").strip()
    where, params = [], []
    if re.match(r"^\d{4}-\d{2}-\d{2}$", start):
        where.append("p.pay_date >= ?")
        params.append(start)
    if re.match(r"^\d{4}-\d{2}-\d{2}$", end):
        where.append("p.pay_date <= ?")
        params.append(end)
    if keyword:
        where.append("(s.name LIKE ? OR s.student_no LIKE ?)")
        params.append(f"%{keyword}%")
        params.append(f"%{keyword}%")
    if course:
        where.append("p.course = ?")
        params.append(course)
    if method:
        where.append("p.method = ?")
        params.append(method)
    cond = ("WHERE " + " AND ".join(where)) if where else ""

    conn = get_db()
    try:
        rows = conn.execute(
            f"""
            SELECT p.id, p.student_id, p.course, p.amount, p.pay_date, p.method, p.remark, p.status,
                   p.unit_price, p.lesson_count, p.discount_percent, p.discount_amount,
                   s.student_no, s.name AS student_name
            FROM payments p JOIN students s ON s.id = p.student_id
            {cond}
            ORDER BY p.pay_date DESC, p.id DESC
            LIMIT 5000
            """,
            params,
        ).fetchall()
        result = _attach_photos(conn, rows, "payment_photos", "payment_id")
    finally:
        conn.close()
    return jsonify(result)


@app.get("/api/payments/export")
@permission_required("payment.export", "当前角色无权导出缴费 Excel")
def export_payments():
    """导出缴费查询结果为 Excel"""
    start = request.args.get("start", "").strip()
    end = request.args.get("end", "").strip()
    keyword = request.args.get("keyword", "").strip()
    course = request.args.get("course", "").strip()
    method = request.args.get("method", "").strip()
    where, params = [], []
    if re.match(r"^\d{4}-\d{2}-\d{2}$", start):
        where.append("p.pay_date >= ?")
        params.append(start)
    if re.match(r"^\d{4}-\d{2}-\d{2}$", end):
        where.append("p.pay_date <= ?")
        params.append(end)
    if keyword:
        where.append("(s.name LIKE ? OR s.student_no LIKE ?)")
        params.append(f"%{keyword}%")
        params.append(f"%{keyword}%")
    if course:
        where.append("p.course = ?")
        params.append(course)
    if method:
        where.append("p.method = ?")
        params.append(method)
    cond = ("WHERE " + " AND ".join(where)) if where else ""

    conn = get_db()
    try:
        rows = conn.execute(
            f"""
            SELECT p.pay_date, s.student_no, s.name AS student_name, p.course,
                   p.unit_price, p.lesson_count,
                   p.amount, p.method, p.remark
            FROM payments p JOIN students s ON s.id = p.student_id
            {cond}
            ORDER BY p.pay_date DESC, p.id DESC
            LIMIT 5000
            """,
            params,
        ).fetchall()
    finally:
        conn.close()

    # ---- 生成 Excel ----
    wb = Workbook()
    ws = wb.active
    ws.title = "缴费查询"

    # 样式
    head_font = Font(name="等线", bold=True, color="FFFFFF")
    head_fill = PatternFill("solid", fgColor="4F46E5")
    title_font = Font(name="等线", size=14, bold=True)
    sub_font = Font(name="等线", size=9, color="808080")
    cell_font = Font(name="等线")
    money_font = Font(name="等线", color="C00000")
    center = Alignment(horizontal="center", vertical="center")
    left = Alignment(horizontal="left", vertical="center", wrap_text=True)
    border = Border(
        left=Side(style="thin", color="D0D0D0"),
        right=Side(style="thin", color="D0D0D0"),
        top=Side(style="thin", color="D0D0D0"),
        bottom=Side(style="thin", color="D0D0D0"),
    )
    sum_fill = PatternFill("solid", fgColor="FFF2CC")

    headers = ["缴费日期", "学号", "学生姓名", "科目", "科目单价(元)", "课程节数",
               "缴费金额(元)", "缴费方式", "备注"]
    widths = [12, 10, 12, 16, 12, 10, 14, 10, 30]

    # 标题
    ws.merge_cells("A1:I1")
    ws["A1"] = "缴费查询明细"
    ws["A1"].font = title_font
    ws["A1"].alignment = center

    # 筛选条件说明
    ws.merge_cells("A2:I2")
    cond_parts = []
    if start:
        cond_parts.append(f"开始: {start}")
    if end:
        cond_parts.append(f"结束: {end}")
    if keyword:
        cond_parts.append(f"学生: {keyword}")
    if course:
        cond_parts.append(f"科目: {course}")
    if method:
        cond_parts.append(f"方式: {method}")
    ws["A2"] = "筛选条件：" + ("；".join(cond_parts) if cond_parts else "全部记录")
    ws["A2"].font = sub_font

    # 表头
    head_row = 4
    for i, (h, w) in enumerate(zip(headers, widths), start=1):
        c = ws.cell(row=head_row, column=i, value=h)
        c.font = head_font
        c.fill = head_fill
        c.alignment = center
        c.border = border
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A5"

    # 数据行
    row = head_row + 1
    total_amount = 0
    for r in rows:
        vals = [
            r["pay_date"] or "",
            r["student_no"] or "",
            r["student_name"] or "",
            r["course"] or "",
            r["unit_price"] or 0,
            r["lesson_count"] or 0,
            r["amount"] or 0,
            r["method"] or "",
            r["remark"] or "",
        ]
        for col, v in enumerate(vals, start=1):
            c = ws.cell(row=row, column=col, value=v)
            c.font = cell_font
            c.border = border
            c.alignment = center if col in (1, 2, 3, 6, 8) else left
            if col in (5, 7):
                c.number_format = '#,##0.00'
            if col == 7:
                c.font = money_font
        total_amount += r["amount"] or 0
        row += 1

    # 汇总行
    row += 1
    ws.cell(row=row, column=1, value="合计").font = Font(name="等线", bold=True)
    ws.cell(row=row, column=1).fill = sum_fill
    ws.cell(row=row, column=6, value="总金额").font = Font(name="等线", bold=True)
    ws.cell(row=row, column=6).fill = sum_fill
    amt_cell = ws.cell(row=row, column=7, value=total_amount)
    amt_cell.font = Font(name="等线", bold=True, color="C00000", size=13)
    amt_cell.fill = sum_fill
    amt_cell.number_format = '#,##0.00'
    ws.cell(row=row, column=8, value=f"共 {len(rows)} 笔").font = sub_font
    ws.cell(row=row, column=8).fill = sum_fill
    for col in range(1, 10):
        ws.cell(row=row, column=col).border = border

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return send_file(
        buf,
        as_attachment=True,
        download_name=datetime.now().strftime("缴费查询_%Y%m%d_%H%M%S.xlsx"),
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


def _payment_payload(d):
    """解析并校验缴费数据，返回 (data, error)"""
    student_id = d.get("student_id")
    if not isinstance(student_id, int):
        return None, "请选择缴费学生"
    amount = _to_float(d.get("amount"))
    if amount is None or amount <= 0:
        return None, "请输入正确的缴费金额"
    pay_date = str(d.get("pay_date") or "").strip()
    if not re.match(r"^\d{4}-\d{2}-\d{2}$", pay_date):
        pay_date = datetime.now().strftime("%Y-%m-%d")
    method = str(d.get("method") or "").strip()
    if method not in PAY_METHODS:
        method = "现金"
    return {
        "student_id": student_id,
        "course": str(d.get("course") or "").strip(),
        "amount": amount,
        "unit_price": _to_float(d.get("unit_price")) or 0,
        "lesson_count": int(d.get("lesson_count") or 0),
        "discount_percent": _to_float(d.get("discount_percent")) or 100,
        "discount_amount": _to_float(d.get("discount_amount")) or 0,
        "pay_date": pay_date,
        "method": method,
        "remark": str(d.get("remark") or "").strip()[:100],
        "status": str(d.get("status") or "已确认").strip(),
    }, None


@app.post("/api/payments")
@permission_required("payment.write", "当前角色无权新增缴费记录")
def create_payment():
    """新建缴费记录"""
    data, error = _payment_payload(request.get_json(silent=True) or {})
    if error:
        return jsonify({"error": error}), 400
    conn = get_db()
    try:
        if not conn.execute(
            "SELECT 1 FROM students WHERE id = ?", (data["student_id"],)
        ).fetchone():
            return jsonify({"error": "学生档案不存在"}), 404
        cur = conn.execute(
            """
            INSERT INTO payments
                (student_id, course, amount, unit_price, lesson_count, discount_percent, discount_amount,
                 pay_date, method, remark, status, created_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (data["student_id"], data["course"], data["amount"],
             data["unit_price"], data["lesson_count"], data["discount_percent"], data["discount_amount"],
             data["pay_date"], data["method"], data["remark"], data["status"],
             datetime.now().strftime("%Y-%m-%d %H:%M:%S")),
        )
        conn.commit()
        payment_id = cur.lastrowid
    finally:
        conn.close()
    return jsonify({"id": payment_id, "message": "缴费记录已保存"}), 201


@app.put("/api/payments/<int:payment_id>")
@permission_required("payment.write", "当前角色无权修改缴费记录")
def update_payment(payment_id):
    """修改缴费记录"""
    data, error = _payment_payload(request.get_json(silent=True) or {})
    if error:
        return jsonify({"error": error}), 400
    conn = get_db()
    try:
        if not conn.execute(
            "SELECT 1 FROM students WHERE id = ?", (data["student_id"],)
        ).fetchone():
            return jsonify({"error": "学生档案不存在"}), 404
        cur = conn.execute(
            """
            UPDATE payments SET
                student_id = ?, course = ?, amount = ?, unit_price = ?, lesson_count = ?,
                discount_percent = ?, discount_amount = ?, pay_date = ?, method = ?, remark = ?, status = ?
            WHERE id = ?
            """,
            (data["student_id"], data["course"], data["amount"],
             data["unit_price"], data["lesson_count"], data["discount_percent"], data["discount_amount"],
             data["pay_date"], data["method"], data["remark"], data["status"], payment_id),
        )
        conn.commit()
        if cur.rowcount == 0:
            return jsonify({"error": "缴费记录不存在"}), 404
    finally:
        conn.close()
    return jsonify({"message": "修改成功"})


@app.delete("/api/payments/<int:payment_id>")
@permission_required("payment.delete", "教务不可删除缴费记录，删除需由管理员或校长操作")
def delete_payment(payment_id):
    """删除缴费记录"""
    conn = get_db()
    try:
        cur = conn.execute("DELETE FROM payments WHERE id = ?", (payment_id,))
        conn.execute("DELETE FROM payment_photos WHERE payment_id = ?", (payment_id,))
        conn.commit()
        if cur.rowcount == 0:
            return jsonify({"error": "缴费记录不存在"}), 404
    finally:
        conn.close()
    return jsonify({"message": "删除成功"})


# ---------- 缴费证明图片 ----------
MAX_PAY_PHOTOS = 5   # 每笔缴费最多证明图片数


def _attach_photos(conn, rows, table, key):
    """为记录列表批量附加证明图片信息（避免逐条 N+1）

    table: payment_photos / refund_photos；key: payment_id / refund_id
    """
    result = []
    ids = [r["id"] for r in rows]
    photos_map = {}
    if ids:
        placeholders = ",".join("?" * len(ids))
        for p in conn.execute(
            f"SELECT id, {key}, filename, original_name FROM {table} "
            f"WHERE {key} IN ({placeholders}) ORDER BY id",
            ids,
        ).fetchall():
            photos_map.setdefault(p[key], []).append(dict(p))
    for r in rows:
        d = dict(r)
        d["photos"] = photos_map.get(d["id"], [])
        result.append(d)
    return result


@app.post("/api/payments/<int:payment_id>/photos")
def upload_payment_photos(payment_id):
    """上传缴费证明图片（每笔缴费最多 5 张，jpg/png/gif/webp，单张 5MB 内）"""
    files = [f for f in request.files.getlist("photos") if f and f.filename]
    if not files:
        return jsonify({"error": "未选择任何图片"}), 400
    conn = get_db()
    try:
        if not conn.execute(
            "SELECT 1 FROM payments WHERE id = ?", (payment_id,)
        ).fetchone():
            return jsonify({"error": "缴费记录不存在"}), 404
        existing = conn.execute(
            "SELECT COUNT(*) AS c FROM payment_photos WHERE payment_id = ?", (payment_id,)
        ).fetchone()["c"]
        if existing + len(files) > MAX_PAY_PHOTOS:
            return jsonify(
                {"error": f"缴费证明最多 {MAX_PAY_PHOTOS} 张（已有 {existing} 张，本次选择 {len(files)} 张）"}
            ), 400

        # 先整体校验（格式/大小），全部通过后再统一保存
        blobs = []
        for f in files:
            ext = os.path.splitext(f.filename)[1].lower()
            if ext not in ALLOWED_PHOTO_EXT:
                return jsonify(
                    {"error": f"不支持的图片格式：{f.filename}（仅支持 jpg/png/gif/webp）"}
                ), 400
            blob = f.read()
            if len(blob) > MAX_PHOTO_SIZE:
                return jsonify(
                    {"error": f"图片过大：{f.filename}（单张不超过 5MB）"}
                ), 400
            blobs.append((blob, ext, os.path.basename(f.filename).replace(" ", "_") or "photo"))

        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        for blob, ext, original in blobs:
            stored = f"pay{payment_id}_{uuid.uuid4().hex[:12]}{ext}"
            conn.execute(
                "INSERT INTO payment_photos (payment_id, filename, original_name, size, uploaded_at, data) "
                "VALUES (?, ?, ?, ?, ?, ?)",
                (payment_id, stored, original, len(blob), now, blob),
            )
        conn.commit()
    finally:
        conn.close()
    return jsonify({"message": f"成功上传 {len(blobs)} 张缴费证明"}), 201


@app.delete("/api/payment-photos/<int:photo_id>")
def delete_payment_photo(photo_id):
    """删除一张缴费证明图片"""
    conn = get_db()
    try:
        cur = conn.execute("DELETE FROM payment_photos WHERE id = ?", (photo_id,))
        conn.commit()
        if cur.rowcount == 0:
            return jsonify({"error": "图片不存在"}), 404
    finally:
        conn.close()
    return jsonify({"message": "删除成功"})


@app.get("/uploads/payment-photos/<path:filename>")
def get_payment_photo(filename):
    """查看缴费证明图片；带 ?download=1 时以原始文件名下载（只读角色不可下载）"""
    if _viewer_download_denied():
        return jsonify({"error": "只读角色无权下载缴费证明"}), 403
    conn = get_db()
    try:
        row = conn.execute(
            "SELECT data, original_name FROM payment_photos WHERE filename = ?", (filename,)
        ).fetchone()
    finally:
        conn.close()
    if not row or not row["data"]:
        abort(404)
    download_name = row["original_name"] or filename
    return send_file(
        io.BytesIO(row["data"]),
        download_name=download_name,
        as_attachment=bool(request.args.get("download")),
    )


# ---------- 科目管理 ----------
# 科目每年固定课次：一年年费 ÷ 96 节 = 每节课单价（供缴费等模块使用）
SUBJECT_LESSONS_PER_YEAR = 96


def _subject_payload():
    """读取并校验科目表单字段（名称 + 一年年费；课单价由年费÷96自动计算）"""
    body = request.get_json(silent=True) or {}
    name = str(body.get("name", "")).strip()
    if not name:
        return None, "请填写科目名称"
    if len(name) > 50:
        return None, "科目名称不能超过 50 个字"
    year_price = _to_float(body.get("year_price"))
    if year_price is None:
        return None, "一年年费需为不小于 0 的数字"
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    return {
        "name": name,
        "price": round(year_price, 2),
        "lesson_count": SUBJECT_LESSONS_PER_YEAR,
        "lesson_price": round(year_price / SUBJECT_LESSONS_PER_YEAR, 2),
        "updated_at": now,
    }, None


@app.route("/subjects")
@login_required
def subjects_page():
    """科目管理页：科目档案仅维护科目名称与每节课课单价"""
    return render_template("subjects.html", user=current_user())


@app.get("/api/subjects")
def list_subjects():
    """科目列表 + 统计（科目数 / 平均课单价）"""
    conn = get_db()
    try:
        subjects = [dict(r) for r in conn.execute(
            """
            SELECT id, name, price, lesson_count, lesson_price, created_at, updated_at
            FROM subjects ORDER BY id DESC
            """
        ).fetchall()]
        total = conn.execute(
            """
            SELECT COUNT(*) AS n,
                   COALESCE(AVG(lesson_price), 0) AS avg_price
            FROM subjects
            """
        ).fetchone()
    finally:
        conn.close()
    return jsonify({
        "subjects": subjects,
        "stats": {
            "count": total["n"],
            "avg_price": round(total["avg_price"], 2),
        },
    })


@app.post("/api/subjects")
@permission_required("subject.write", "当前角色无权新增科目")
def add_subject():
    """新增科目"""
    data, err = _subject_payload()
    if err:
        return jsonify({"error": err}), 400
    conn = get_db()
    try:
        cur = conn.execute(
            """
            INSERT INTO subjects (name, price, lesson_count, lesson_price, created_at, updated_at)
            VALUES (?, ?, ?, ?, ?, ?)
            """,
            (data["name"], data["price"], data["lesson_count"], data["lesson_price"],
             data["updated_at"], data["updated_at"]),
        )
        conn.commit()
        sid = cur.lastrowid
    finally:
        conn.close()
    return jsonify({"id": sid, "message": "新增成功"}), 201


@app.put("/api/subjects/<int:subject_id>")
@permission_required("subject.write", "当前角色无权修改科目")
def update_subject(subject_id):
    """修改科目（名称 / 课单价）"""
    data, err = _subject_payload()
    if err:
        return jsonify({"error": err}), 400
    conn = get_db()
    try:
        cur = conn.execute(
            """
            UPDATE subjects
               SET name = ?, price = ?, lesson_count = ?, lesson_price = ?, updated_at = ?
             WHERE id = ?
            """,
            (data["name"], data["price"], data["lesson_count"], data["lesson_price"],
             data["updated_at"], subject_id),
        )
        conn.commit()
        if cur.rowcount == 0:
            return jsonify({"error": "科目不存在"}), 404
    finally:
        conn.close()
    return jsonify({"message": "修改成功"})


@app.delete("/api/subjects/<int:subject_id>")
@permission_required("subject.delete", "教务不可删除所创建的科目，删除需由管理员或校长操作")
def delete_subject(subject_id):
    """删除科目"""
    conn = get_db()
    try:
        cur = conn.execute("DELETE FROM subjects WHERE id = ?", (subject_id,))
        conn.commit()
        if cur.rowcount == 0:
            return jsonify({"error": "科目不存在"}), 404
    finally:
        conn.close()
    return jsonify({"message": "删除成功"})


# ---------- 报表统计 ----------
@app.route("/reports")
@permission_required("report.view", "仅管理员与校长可查看数据报表")
def reports_page():
    """报表统计页（经营数据看板）"""
    conn = get_db()
    try:
        now = datetime.now()
        month_prefix = now.strftime("%Y-%m")
        today_str = now.strftime("%Y-%m-%d")
        month_ago = (now - timedelta(days=29)).strftime("%Y-%m-%d")

        # 学生
        stu_total = conn.execute("SELECT COUNT(*) AS c FROM students").fetchone()["c"]
        stu_male = conn.execute(
            "SELECT COUNT(*) AS c FROM students WHERE gender = '男'").fetchone()["c"]
        stu_female = stu_total - stu_male
        stu_month = conn.execute(
            "SELECT COUNT(*) AS c FROM students WHERE created_at LIKE ?",
            (month_prefix + "%",)).fetchone()["c"]
        course_dist = [dict(r) for r in conn.execute(
            """
            SELECT COALESCE(NULLIF(TRIM(course), ''), '未填写') AS k, COUNT(*) AS c
            FROM students GROUP BY k ORDER BY c DESC LIMIT 8
            """)]

        # 课程
        course_total = conn.execute("SELECT COUNT(*) AS c FROM courses").fetchone()["c"]
        course_active = conn.execute(
            "SELECT COUNT(*) AS c FROM courses WHERE status = '开课中'").fetchone()["c"]

        # 教师
        tea_total = conn.execute("SELECT COUNT(*) AS c FROM teachers").fetchone()["c"]
        tea_active = conn.execute(
            "SELECT COUNT(*) AS c FROM teachers WHERE status = '在职'").fetchone()["c"]

        # 缴费
        pay_count = conn.execute("SELECT COUNT(*) AS c FROM payments").fetchone()["c"]
        pay_total = conn.execute(
            "SELECT COALESCE(SUM(amount), 0) AS s FROM payments").fetchone()["s"]
        pay_month = conn.execute(
            "SELECT COALESCE(SUM(amount), 0) AS s FROM payments WHERE pay_date LIKE ?",
            (month_prefix + "%",)).fetchone()["s"]
        pay_today = conn.execute(
            "SELECT COALESCE(SUM(amount), 0) AS s FROM payments WHERE pay_date = ?",
            (today_str,)).fetchone()["s"]
        pay_by_course = [dict(r) for r in conn.execute(
            """
            SELECT COALESCE(NULLIF(TRIM(course), ''), '未填写') AS k,
                   COUNT(*) AS n, COALESCE(SUM(amount), 0) AS s
            FROM payments GROUP BY k ORDER BY s DESC LIMIT 5
            """)]
        pay_by_method = [dict(r) for r in conn.execute(
            """
            SELECT COALESCE(NULLIF(TRIM(method), ''), '未填写') AS k,
                   COUNT(*) AS n, COALESCE(SUM(amount), 0) AS s
            FROM payments GROUP BY k ORDER BY s DESC
            """)]

        # 近6个月缴费趋势
        pay_monthly = []
        for i in range(5, -1, -1):
            m = (now.replace(day=1) - timedelta(days=i * 30)).strftime("%Y-%m")
            row = conn.execute(
                "SELECT COALESCE(SUM(amount), 0) AS s, COUNT(*) AS n FROM payments WHERE pay_date LIKE ?",
                (m + "%",)).fetchone()
            pay_monthly.append({"m": m, "s": row["s"], "n": row["n"]})

        # 考勤
        att = {k: 0 for k in ATT_STATUS}
        for r in conn.execute("SELECT status, COUNT(*) AS c FROM attendance GROUP BY status"):
            if r["status"] in att:
                att[r["status"]] = r["c"]
        att_total = sum(att.values())
        att_rate = round((att["present"] + att["late"]) / att_total * 100) if att_total else 0

        att30 = {k: 0 for k in ATT_STATUS}
        for r in conn.execute(
            "SELECT status, COUNT(*) AS c FROM attendance WHERE date >= ? GROUP BY status",
            (month_ago,),
        ):
            if r["status"] in att30:
                att30[r["status"]] = r["c"]
        att30_total = sum(att30.values())
        att30_rate = round((att30["present"] + att30["late"]) / att30_total * 100) if att30_total else 0
    finally:
        conn.close()

    return render_template(
        "reports.html",
        stu_total=stu_total, stu_male=stu_male, stu_female=stu_female, stu_month=stu_month,
        course_dist=course_dist,
        course_total=course_total, course_active=course_active,
        tea_total=tea_total, tea_active=tea_active,
        pay_count=pay_count, pay_total=pay_total, pay_month=pay_month, pay_today=pay_today,
        pay_by_course=pay_by_course, pay_by_method=pay_by_method,
        pay_monthly=pay_monthly,
        att=att, att_total=att_total, att_rate=att_rate,
        att30=att30, att30_total=att30_total, att30_rate=att30_rate,
        month=month_prefix, user=current_user(),
    )


# ---------- 启动 ----------
if __name__ == "__main__":
    init_db()
    port = int(os.environ.get("PORT", 5000))
    ip = get_lan_ip()
    print("=" * 56)
    print("  培训机构管理系统 —— 学生档案")
    print("-" * 56)
    print(f"  本机访问:   http://127.0.0.1:{port}")
    print(f"  局域网访问: http://{ip}:{port}")
    print("-" * 56)
    print("  其他电脑/手机需与本机处于同一局域网，")
    print("  并在防火墙中放行该端口后即可访问。")
    print("=" * 56)
    app.run(host="0.0.0.0", port=port, debug=False)
