# 培训机构管理系统 — 技术方案文档

| 项目 | 说明 |
|------|------|
| 文档版本 | v3.4 |
| 创建日期 | 2026-09-14 |
| 更新日期 | 2026-10-07 |
| 技术栈 | Python Flask + SQLite + 原生 HTML/CSS/JS + Chart.js |
| 代码规模 | `app.py` ~4600 行，16 个 HTML 模板，16 个 CSS 文件，1 个公共 JS |
| 接口规模 | 14 个页面路由 + 72 个 API 路由 |

---

## 1. 系统架构

### 1.1 架构图

```
浏览器 (Chrome/Edge/Safari/移动端)
  │  HTML(Jinja2渲染) + CSS(static) + JS(common.js + 内联fetch交互)
  │  HTTP: JSON / multipart, 带Session Cookie
  ▼
Flask 应用 (app.py 单文件)
  ├── 认证层: 登录/Session/角色权限装饰器/账号锁定
  ├── 路由层: 14个页面路由 + 71个API路由
  ├── 业务层: 校验/清洗/编号生成/附件元信息/课时核算/操作日志
  ├── 数据层: sqlite3原生驱动 + Row Factory + WAL模式
  └── 工具层: openpyxl Excel导出 / 密码加密 / 日志 / 备份与恢复
  │
  ▼
SQLite (students.db 单文件, WAL模式, 含BLOB存储照片和附件)
  ├── backups/data/ 目录 — 一键备份 / 下载 / 恢复（恢复前自动留存快照）
  └── operation_logs 表 — 操作审计
```

### 1.2 技术选型

| 组件 | 选型 | 理由 |
|------|------|------|
| 后端 | Flask 3.0+ | 轻量，单文件承载全部业务，适合内网部署 |
| 数据库 | SQLite (WAL模式) | 零配置、单文件、支持BLOB、无需独立服务，WAL提升并发读 |
| 前端 | 原生HTML+CSS+JS + Chart.js | 无构建步骤，无Node依赖，Chart.js提供图表可视化 |
| Excel | openpyxl 3.1+ | 纯Python，精确控制样式/合并/冻结 |
| 密码加密 | werkzeug.security | PBKDF2-SHA256，Flask内置依赖无需额外安装 |
| 会话 | Flask Session | 服务端签名Cookie，零配置 |

### 1.3 前后端交互

- **页面**：Jinja2 服务端渲染，非 SPA
- **数据**：`fetch()` 调 RESTful API，返回 JSON，带 Session Cookie
- **上传**：`multipart/form-data`
- **下载**：`send_file()` 从 BLOB → `io.BytesIO` → 响应流
- **前端公共组件**：`common.js` 提供 Toast 通知、自定义模态框、fetch 封装、表单校验
- **CSS 变量**：`:root` 定义全局主题色/间距/圆角，所有页面统一引用

---

## 2. 数据库设计

### 2.1 表清单

| 表名 | 说明 | 关键约束 |
|------|------|----------|
| `users` | 系统用户 | `username` UNIQUE, password_hash |
| `operation_logs` | 操作审计日志 | user_id, action, module, ip |
| `students` | 学生档案 | `student_no` UNIQUE |
| `guardians` | 监护人 | 一对多关联 students |
| `photos` | 学生照片 | `data` BLOB 存二进制 |
| `resume_items` | 简历条目 | type: 6种类型 |
| `attachments` | 简历附件 | `data` BLOB 存二进制 |
| `attendance` | 考勤记录 | `UNIQUE(student_id, date)`；含 start_time/end_time 上课时间段与 source 来源（空=正常登记，makeup=补考勤） |
| `courses` | 课程 | status: 开课中/已结课, teacher_id 外键 |
| `enrollments` | 选课记录 | `UNIQUE(course_id, student_id)` |
| `teachers` | 教师 | `teacher_no` UNIQUE |
| `payments` | 缴费记录 | amount NOT NULL, status 字段 |
| `payment_photos` | 缴费证明图片 | `data` BLOB 存二进制，payment_id 索引 |
| `refunds` | 退费记录 | payment_id 关联原缴费，lesson_count 退课时数 |
| `refund_photos` | 退费证明图片 | `data` BLOB 存二进制，refund_id 索引 |
| `subjects` | 科目 | lesson_price 核心字段；year_price / lessons_per_year 年费信息 |
| `system_settings` | 系统设置 | key-value 结构 |
| `counters` | 编号计数器 | student_seq / teacher_seq |

### 2.2 ER 关系

```
users (独立，认证与权限)
operation_logs (独立，审计日志)
students ──< guardians
students ──< photos
students ──< resume_items ──< attachments
students ──< attendance
students ──< payments ──< refunds
students ──< enrollments >── courses
students ──< payments ──< payment_photos (缴费证明 BLOB)
students ──< refunds  ──< refund_photos (退费证明 BLOB)
teachers ──< courses (通过 teacher_id 外键关联)
subjects (独立，被 payments.course 按名称引用)
system_settings (独立，key-value 配置)
counters (独立，存储自增序列)
```

### 2.3 数据库迁移机制

`init_db()` 启动时自动执行，实现向后兼容：

1. `PRAGMA journal_mode=WAL`（启用 WAL 模式）
2. `PRAGMA foreign_keys=ON`（启用外键约束）
3. `CREATE TABLE IF NOT EXISTS`（16张表）
4. `PRAGMA table_info` 检测列，缺失则 `ALTER TABLE ADD COLUMN`
5. 旧数据迁移：编号回填、监护人迁移、磁盘文件入BLOB
6. `CREATE INDEX IF NOT EXISTS`（关键查询字段索引）
7. 初始化默认管理员账号（首次启动）

| 迁移项 | 说明 |
|--------|------|
| `photos.sort` | 旧表无排序列，自动添加 |
| `payments.unit_price/lesson_count` | 旧表无计价字段，自动添加 |
| `students.student_no/grade/id_card` | 旧表无这些列，自动添加 |
| `students.student_no` 回填 | 已有学生按id顺序生成编号 |
| `teachers.teacher_no` 回填 | 已有教师按id顺序生成工号 |
| `courses.start_date/end_date` | 自动添加并回填 |
| `attendance.course_name/lesson/teacher` | 自动添加 |
| `attendance.start_time/end_time/source` | v3.1 新增上课时间段（HH:MM）与来源列（空=正常登记，makeup=补考勤），自动添加 |
| `refunds.lesson_count` | v3.2 新增退课时数（默认0），退费时同步扣减学员剩余课时，自动添加 |
| `payment_photos` / `refund_photos` 表 | v3.3 新建（缴费/退费证明图片 BLOB），含 payment_id / refund_id 索引，建表自动完成 |
| `subjects.year_price/lessons_per_year` | 年费与年课时数字段，自动添加 |
| `courses.teacher_id` | 新增外键字段，关联 teachers.id |
| `payments.status` | 新增缴费状态字段（已确认/待确认） |
| `users` 表 | 新建，含默认管理员 admin/admin123 |
| `operation_logs` 表 | 新建，审计日志 |
| `refunds` 表 | 新建，退费记录 |
| `system_settings` 表 | 新建，系统配置 |
| 监护人迁移 | students表旧字段 → guardians表 |
| 照片BLOB迁移 | 磁盘文件 → photos.data BLOB |
| 附件BLOB迁移 | 磁盘文件 → attachments.data BLOB |

---

## 3. 后端核心设计

### 3.1 用户认证与权限

**密码加密**：使用 `werkzeug.security.generate_password_hash` (PBKDF2-SHA256) 加密存储，`check_password_hash` 验证。

**Session 管理**：
- Flask Session 签名 Cookie，`SECRET_KEY` 随机生成或环境变量配置
- 登录成功后 `session['user_id']` 和 `session['role']` 写入
- `before_request` 钩子检查 Session 有效性，未登录重定向 `/login`
- Session 有效期 8 小时，每次请求刷新

**角色权限装饰器**（v3.4 权限点化重构）：

5 个角色：admin（管理员）/ principal（校长）/ staff（教务）/ teacher（教师）/ viewer（只读），前两者为唯一角色（全系统各一个）。全部权限收敛为权限点集中定义：

```python
PERMISSION_MATRIX = {
    "attendance.save":   ATTENDANCE_SAVE_ROLES,    # admin/principal/staff/teacher
    "attendance.batch":  ATTENDANCE_BATCH_ROLES,   # admin/principal/staff
    "attendance.delete": ATTENDANCE_DELETE_ROLES,  # admin/principal/staff
    "attendance.export": ATTENDANCE_EXPORT_ROLES,  # admin/principal/staff
    "course.write":      COURSE_WRITE_ROLES,       # admin/principal/staff/teacher
    "settings.logs":     SETTINGS_LOGS_ROLES,      # admin/principal/staff/teacher
    "report.view":       REPORT_VIEW_ROLES,        # admin/principal
    "import.all":        IMPORT_ROLES,             # admin/principal/staff
    "settings.manage":   SUPER_ROLES,              # admin/principal
    # ... student/resume/payment/refund/subject/teacher/lesson 等同理
}

def can(action)                                  # 模板与后端共用：session.role 是否在允许列表
def _role_required(roles, message)               # 参数化装饰器工厂
def permission_required(action, message)         # 按权限点校验
def login_required(f)                            # 检查是否登录
def student_write_required(f) / resume_write_required(f) / super_required(f)  # 语义化包装
```

前端模板通过全局 `can(action)` 裁剪功能入口（无权限的功能不渲染），后端接口用装饰器同步拦截（403），前端隐藏不是唯一防线。

**登录日志关联修复（v3.4）**：`api_login` 原在 Session 建立前写登录日志，导致 `user_id=NULL`、username=anonymous；已将 `log_operation` 移到 session 设置之后，登录日志正确关联用户。

**账号锁定**：
- `users` 表增加 `failed_attempts` 和 `locked_until` 字段
- 连续 5 次密码错误设置 `locked_until = now + 15min`
- 登录时检查锁定状态

**默认管理员**：`init_db()` 时检查 `users` 表是否为空，空则创建 `admin/admin123`（`last_login=NULL`，首次登录强制改密码）。

### 3.2 操作审计日志

| 函数 | 用途 |
|------|------|
| `log_operation(conn, action, module, target, detail)` | 记录操作日志 |

- 在 `before_request` 中获取当前用户和 IP
- 关键操作（创建/修改/删除/登录/导出/备份）调用 `log_operation()`
- 日志记录字段：user_id, username, action, module, target, detail, ip, created_at
- 日志查询支持按日期范围、用户名、操作类型、模块筛选

### 3.3 编号生成

**学号** `a00001` ~ `z99999`：counters 表 `student_seq` 自增，字母 a-z + 五位数字，容量 2,599,974。

**教师工号** `a0001` ~ `z9999`：同理，四位数字，容量 259,974。

### 3.4 数据校验

| 校验项 | 规则 |
|--------|------|
| 学生姓名 | 必填 |
| 身份证号 | 18位，末位可为X（正则 `^\d{17}[\dXx]$`） |
| 性别 | 非"男""女"默认"男" |
| 监护人 | 姓名/电话/微信全空的组自动丢弃 |
| 课程名称 | 必填 |
| 教师姓名 | 必填 |
| 缴费金额 | 必填，必须 > 0 |
| 缴费状态 | 默认"已确认"，可选"已确认"/"待确认" |
| 退费金额 | 必填，必须 > 0，不超过原缴费金额 |
| 科目名称 | 必填，≤50字 |
| 课单价 | 必填，≥0 的数字 |
| 用户名 | 必填，≥3字符，唯一 |
| 密码 | 必填，≥6字符 |

### 3.5 级联删除

删除学生时一次性清理关联数据：photos、attachments、resume_items、attendance、payment_photos（按学生缴费）、payments、refund_photos（按学生退费）、refunds、enrollments、guardians。

删除单条缴费记录时同步删除其 `payment_photos`；删除单条退费记录时同步删除其 `refund_photos`。

**删除用户级联删除其操作日志（v3.4）**：`DELETE /api/users/<id>` 删除账户的同一事务内执行 `DELETE FROM operation_logs WHERE user_id=? OR username=?`（双条件兼容历史无 user_id 的日志）；「删除用户」操作本身记录在删除者名下，保证审计可追溯。

### 3.5.1 证明图片附加查询（v3.3）

**通用函数** `_attach_photos(conn, rows, table, key)`：为缴费/退费列表批量附加图片信息，单次 `IN (...)` 查询构建 `payment_id → [photos]` 映射，避免逐条 N+1：

- 缴费：`_attach_photos(conn, rows, "payment_photos", "payment_id")`
- 退费：`_attach_photos(conn, rows, "refund_photos", "refund_id")`

查询显式排除 `data` BLOB 列（只取 id/filename/original_name），列表接口返回轻量元信息；图片本体经 `/uploads/payment-photos/<filename>` 与 `/uploads/refund-photos/<filename>` 按需加载，`?download=1` 以原始文件名下载。

### 3.6 考勤 UPSERT 与补考勤

利用 SQLite `ON CONFLICT(student_id, date) DO UPDATE SET ...`，同一学生同日重复保存自动覆盖（覆盖 status/remark/course_name/lesson/teacher/start_time/end_time/source）。

**字段校验**：
- `start_time`/`end_time`：截取前 5 位，须匹配 `^\d{2}:\d{2}$`，否则置空
- `source`：仅允许 `""`（正常登记）或 `"makeup"`（补考勤），其他值置空
- `student_id` 必须存在于 students 表，否则跳过该条
- 单次保存上限 1000 条

**补考勤**：前端「补考勤」选项卡以 `source=makeup` 标记写入，复用同一保存接口；查询时通过 `_history_filter` 的 source 参数区分（`/api/attendance/history?source=makeup` 只返回补考勤记录）。

### 3.6.2 考勤记录批量修改（v3.4 新增）

`POST /api/attendance/batch`（权限 `attendance.batch` = admin/principal/staff）：按记录 id 批量更新状态与/或备注。

- 入参：`{"ids": [...], "status": "present|late|leave|absent|", "remark": "文本|null"}`
- 语义：`status` 传空 = 不改状态；`remark` 不传（null）= 保持原备注，传字符串（含空串）= 覆盖
- 校验：ids 去重过滤非法值、单次上限 1000 条、status 必须合法、状态与备注全空返回 400
- 实现：单条 `UPDATE attendance SET status=?, remark=?, updated_at=? WHERE id IN (...?)`，一次事务
- 前端：记录查询页汇总栏「修改查询信息状态」按钮（`can('attendance.batch')` 控制渲染），弹窗列出最近一次查询结果（`historyRecords`），默认全选支持全选/反选/清空，保存后自动重新查询；查询结果为空时弹窗仍可打开并提示原因

### 3.6.3 只读角色下载拦截（v3.4 新增）

通用函数 `_viewer_download_denied()`：`session.role == "viewer" && request.args.download` 时为真。四个附件查看路由（学生照片 `/uploads/photos`、缴费证明 `/uploads/payment-photos`、退费证明 `/uploads/refund-photos`、简历附件 `/uploads/files`）在处理 `?download=1` 时检查并返回 403，**在线预览（无 download 参数）不受限**。前端缩略图点击行为按 `PERMS.download` 标志同步裁剪。

考勤登记页对 `PERMS.save == false` 的角色（viewer）：班级卡教师/上课时间/上课日期/本节课名称与学员状态、备注全部渲染为纯文本（下拉与输入框不生成），操作栏（保存/全部出勤/临时学员）不渲染；数据加载函数前置执行并加 try/catch，保证局部渲染异常不影响页面数据展示。

### 3.6.1 课时核算（v3.2 新增）

**统一查询函数** `_lesson_stats_rows(keyword, warn_only, class_id, teacher_id)`：一次 SQL 聚合出学生维度的课时全貌，被课时管理页、首页统计、学生列表三处复用，保证口径唯一。

```sql
SELECT s.id, s.student_no, s.name,
       COALESCE(p_t.t, 0) AS total_lessons,      -- 已购：payments.lesson_count 求和
       COALESCE(r_t.t, 0) AS refunded_lessons,   -- 已退：refunds.lesson_count 求和
       COALESCE(a_t.c, 0) AS used_lessons,       -- 已消耗：attendance 出勤+迟到 计数
       COALESCE(p_t.t,0) - COALESCE(r_t.t,0) - COALESCE(a_t.c,0) AS remaining_lessons
FROM students s
LEFT JOIN (SELECT student_id, SUM(lesson_count) AS t FROM payments GROUP BY student_id) p_t ON ...
LEFT JOIN (SELECT student_id, SUM(lesson_count) AS t FROM refunds  GROUP BY student_id) r_t ON ...
LEFT JOIN (SELECT student_id, COUNT(*) AS c FROM attendance
           WHERE status IN ('present','late') GROUP BY student_id) a_t ON ...
ORDER BY remaining_lessons ASC   -- 剩余最少的排最前，便于优先跟进续费
LIMIT 5000
```

**派生指标**（Python 侧计算）：
- `use_rate = used_lessons * 100 / (total_lessons - refunded_lessons)`，分母 ≤ 0 时为 0
- `warn = remaining_lessons <= 0`（含未购课但已上课的场景）

**筛选条件**：
| 参数 | 实现 |
|------|------|
| `keyword` | `s.name LIKE ? OR s.student_no LIKE ?` |
| `warn_only=1` | `remaining_lessons <= 0` |
| `class_id` | `s.id IN (SELECT student_id FROM enrollments WHERE course_id=?)` |
| `teacher_id` | 该教师名下所有班级学员，兼容 `c.teacher_id=?` 与 `c.teacher=(SELECT name FROM teachers WHERE id=?)` 两种关联 |

**退课时数校验**：新建/修改退费时校验 `同缴费记录累计退费金额 ≤ 原缴费金额` 且 `累计退课时数 ≤ 原缴费节数`，超出返回 400 并提示「已退 / 本次 / 原缴费」三项明细。

### 3.7 BLOB 存储流程

**上传**：`request.files` → 格式校验 → 大小校验 → `f.read()` 获取 bytes → `INSERT INTO ... (data) VALUES (?, blob)`

**读取**：`SELECT data FROM photos WHERE filename=?` → `io.BytesIO(data)` → `send_file()`

**关键注意**：查询时显式指定列名排除 `data` BLOB 列，避免 JSON 序列化失败。

### 3.8 工具函数

| 函数 | 用途 |
|------|------|
| `get_db()` | 获取SQLite连接，设置Row Factory |
| `next_student_no(conn)` | 生成下一学号 |
| `next_teacher_no(conn)` | 生成下一教师工号 |
| `clean_payload(data)` | 提取并清理前端提交的10个字段 |
| `extract_guardians(data)` | 提取监护人列表，过滤空组 |
| `validate(payload)` | 基础校验（姓名、身份证号） |
| `get_lan_ip()` | 获取本机局域网IP |
| `safe_remove(path)` | 安全删除文件（带重试） |
| `attach_meta(a)` | 附件附加图标类型与可读大小 |
| `load_resume_sections(conn, sid)` | 按类型分组加载简历条目 |
| `_history_filter(start,end,status,kw,source)` | 构造考勤查询WHERE子句（含 source=makeup 补考勤筛选） |
| `_history_rows_stats(cond,params)` | 考勤明细+各状态统计 |
| `_lesson_stats_rows(kw,warn,class_id,teacher_id)` | **课时核算统一查询**（学生维度已购/已退/已消耗/剩余/使用率/预警，供课时页、首页、学生列表复用） |
| `_attach_photos(conn,rows,table,key)` | **证明图片批量附加**（缴费/退费列表通用，单次 IN 查询避免 N+1） |
| `_to_int(v)` / `_to_float(v)` | 安全类型转换 |
| `_payment_payload(d)` | 缴费数据校验 |
| `_subject_payload()` | 科目数据校验 |
| `log_operation(conn, action, module, target, detail)` | 记录操作审计日志 |
| `current_user()` | 获取当前登录用户（从 Session） |
| `can(action)` | 权限点判断（模板与后端共用 PERMISSION_MATRIX） |
| `permission_required(action, message)` | 装饰器：按权限点校验（v3.4 参数化权限体系） |
| `_role_required(roles, message)` | 装饰器工厂：按角色列表校验 |
| `student_write_required(f)` / `resume_write_required(f)` | 语义化权限装饰器 |
| `_viewer_download_denied()` | 只读角色附件下载拦截判断（download=1 时 403） |
| `login_required(f)` | 装饰器：检查是否登录 |
| `admin_required(f)` | 装饰器：检查是否 admin |
| `list_backups()` | 列出 `backups/data/` 备份（含旧备份迁移） |
| `backup_database()` | 一键备份数据库到 `backups/data/` 目录 |
| `_migrate_legacy_backups()` | 一次性迁移：旧版散落在 `backups/` 根目录的 `.db` 移入 `backups/data/` |
| `get_setting(key, default)` | 读取系统设置项 |

---

## 4. 前端设计

### 4.1 页面列表

| 模板 | 路由 | 功能 |
|------|------|------|
| `login.html` | `/login` | 登录页 |
| `home.html` | `/` | 首页控制台 |
| `index.html` | `/students` | 学生档案列表（分页+筛选+排序） |
| `resume.html` | `/resume/<id>` | 简历查看页（可打印） |
| `resume_edit.html` | `/resume/<id>/edit` | 简历编辑页 |
| `attendance.html` | `/attendance` | 考勤管理（三选项卡：班级卡考勤登记 + 补考勤 + 记录查询） |
| `lessons.html` | `/lessons` | 课时管理（学员剩余课时汇总 + 预警 + 导出） |
| `courses.html` | `/courses` | 课程管理 |
| `teachers.html` | `/teachers` | 教师管理 |
| `payments.html` | `/payments` | 缴费管理（自动计算） |
| `refunds.html` | `/refunds` | 退费管理 |
| `subjects.html` | `/subjects` | 科目管理 |
| `import.html` | `/import` | 数据导入 |
| `reports.html` | `/reports` | 报表统计（图表可视化） |
| `settings.html` | `/settings` | 系统设置（所有角色可进入；标签页按权限：管理员/校长全量，教务/教师为操作日志+修改密码，只读仅修改密码） |

### 4.2 前端模式

- **数据获取**：`fetch()` + `res.json()`，封装在 `common.js` 的 `apiFetch()` 中自动处理 401 跳转
- **DOM渲染**：`innerHTML` + 模板字符串 + `escapeHtml()`
- **文件上传**：`FormData` + `fetch(POST, body=formData)`
- **文件下载**：`window.location.href = exportUrl` 触发浏览器下载
- **CSS**：`common.css` 定义全局 CSS 变量与公共组件，每页独立 CSS 文件
- **JS 公共组件**：`common.js` 提供 Toast 通知、自定义模态框（替代 alert/confirm）、表单校验、防重复提交
- **图表**：Chart.js CDN 引入，报表页渲染柱状图/饼图/折线图/环形图

### 4.3 UI 设计体系（v3.0 · IDE 面板式科技感）

- **单一事实源**：全部设计 token 定义于 `common.css` 的 `:root`，业务页面 CSS 不再重复定义变量；换肤改一处即全局生效
- **变量命名保持兼容**：v2.0 的变量名（`--primary`、`--primary-light` 等）全部保留只改值，另补充 `--ink`（导航墨色）、`--mono`（等宽字体栈）、`--primary-border`、`--danger-bg`、`--money` 等，保证旧引用零断裂
- **批量换肤机制**：10 个业务 CSS 删除重复 `:root` 收敛至公共文件；硬编码 indigo（`#4f46e5`/`#c7d2fe` 等）与页头渐变统一替换为 token 引用；`border-radius` 全局收敛至 4–8px
- **等宽字体策略**：`common.css` 用组合选择器（`.stat-value, .kpi-value, .td-id, .td-amount, .result-stat .num, .bar-value`）统一注入 `var(--mono)`，业务 CSS 未对这些类设置 `font-family`，无需逐页修改即可生效
- **登录页独立深色主题**：`login.css` 深墨背景 + 半透明面板，局部覆盖 `common.css` 的浅色表单样式
- **简历页隔离**：`resume.css` 不引用任何公共变量（朱砂宋体文书打印版式），与换肤体系天然隔离，冻结不改
- **品牌标识**：导航与登录页 logo 统一为代码符号 `</>`（等宽青色），版本号 v3.0

### 4.4 移动端适配设计（v3.3）

**核心原则**：桌面端（>768px）零改动，所有移动体验收敛在 `common.css` 的 `@media (max-width: 768px)` 断点内，改动集中在两个公共文件（`navbar.html` + `common.css`），全站 14 个页面同时生效：

| 机制 | 实现 |
|------|------|
| 汉堡菜单 | `.nav-toggle`（桌面 `display:none`）+ `body.nav-open` 状态类；导航链接 `.nav-links` 在断点内改 `position:fixed` 抽屉（4 列 grid），展开时显示 `.nav-mask` 半透明遮罩；`navbar.html` 内约百字原生 JS 控制开合、点链接/遮罩收起、resize>768 自动复位 |
| 弹窗全屏 | `.modal-overlay.modal-overlay { padding:0; align-items:stretch }`（双类选择器提升优先级，覆盖各页 CSS 后加载的 `.modal-overlay` 样式）；`.modal-overlay .modal, .modal-overlay .modal-box` 全屏 + `100dvh`；兼容 payments 系整页滚动与 settings 系 flex 结构两种弹窗 |
| 表格横滑 | `.table-card / .list-card / .settings-card / .recent-card / .class-table` 断点内 `overflow-x:auto` + `-webkit-overflow-scrolling:touch` |
| 选项卡 | `.settings-tabs / .tabs` 横滑 + 子项 `flex-shrink:0; white-space:nowrap`（防止压缩成竖排） |
| 触控优化 | `.btn-sm`、分页按钮加大；Toast 通栏；`env(safe-area-inset-bottom)` 适配 iPhone 底部 |
| 上传交互 | 缴费/退费证明缩略图下载用隐藏 `<a>` 触发（不 `window.open`，避免 IDE 内置预览接管） |

---

## 5. API 接口设计

### 5.1 规范

- **风格**：RESTful
- **认证**：Session Cookie，所有 API 需登录（除 `/api/auth/login`）
- **权限**：`@login_required` 全局保护；写接口按权限点挂 `@permission_required(...)` 装饰器（与前端 `can()` 共用 PERMISSION_MATRIX），权限不足返回 403
- **请求**：JSON 或 multipart
- **响应**：JSON
- **错误**：HTTP状态码 + `{"error": "message"}`
- **401 未登录**：返回 JSON `{"error": "未登录", "redirect": "/login"}`，前端自动跳转

### 5.2 状态码

| 状态码 | 场景 |
|--------|------|
| 200 | 查询/修改/删除成功 |
| 201 | 创建成功（返回 `{"id":..., "message":...}`） |
| 400 | 参数校验失败 |
| 401 | 未登录或 Session 过期 |
| 403 | 权限不足 |
| 404 | 资源不存在 |

### 5.3 接口汇总（72个）

**用户认证**（7个）：登录/登出/当前用户/改密码 + 用户CRUD（删除用户级联删除其操作日志）

**学生档案**（11个）：CRUD + 照片管理 + 自我评价 + 导入模板/批量导入

**简历**（5个）：条目CRUD + 附件上传/删除

**文件**（2个）：照片查看 + 附件下载（只读角色禁止 ?download=1 附件下载）

**考勤**（6个）：某日查询/批量保存（含补考勤 source=makeup）/**批量修改记录状态/备注（v3.4 新增）**/删除/历史查询（含 source 筛选）/导出Excel

**课时**（3个，v3.2 新增）：筛选选项/学员课时列表+统计/导出Excel

**课程**（5个）：CRUD + 学员名单同步

**教师**（4个）：CRUD

**缴费**（10个）：列表/查询/导出/CRUD + 科目课单价查询 + 证明图片上传/删除/查看（v3.3）

**退费**（8个）：CRUD + 导出Excel + 证明图片上传/删除/查看（v3.3）

**科目**（4个）：CRUD

**系统设置**（8个）：设置读写 + 备份/备份列表/下载/恢复/删除 + 操作日志

---

## 6. 安全设计

| 措施 | 实现 |
|------|------|
| 用户认证 | Session Cookie 签名，所有 API 需登录 |
| 密码加密 | PBKDF2-SHA256（werkzeug.security） |
| 账号锁定 | 连续5次密码错误锁定15分钟 |
| 角色权限 | admin/principal/staff/teacher/viewer 五级；权限点集中管理（PERMISSION_MATRIX），前端 can() 裁剪入口 + 后端 permission_required 拦截（v3.4） |
| 只读下载限制 | viewer 角色对缴费/退费证明、学生照片、简历附件的 `?download=1` 请求返回 403（在线预览不受限） |
| 页面权限裁剪 | 导航/首页模块/页面按钮按 can() 渲染：数据报表仅 report.view、数据导入仅 import.all、系统设置页内按角色裁剪标签页 |
| 操作审计 | operation_logs 表记录关键操作（含备份/恢复/下载） |
| 恢复防护 | 恢复前校验备份为有效 SQLite 库且含 students/users 表；恢复前自动留存当前库快照；均需 admin 权限 |
| 路径穿越防护 | 备份下载/恢复/删除接口校验文件名以 `.db` 结尾且不含 `/`、`\` |
| 身份证脱敏 | 简历页：前6位 + 8星号 + 后4位 |
| SQL注入防护 | 全部参数化查询（`?` 占位符） |
| XSS防护 | 前端 `escapeHtml()` 转义 |
| CSRF防护 | Session Cookie + SameSite 属性 |
| 文件类型限制 | 照片：jpg/png/gif/webp；附件：不限 |
| 文件大小限制 | 照片5MB/张，附件50MB/个，缴费/退费证明5MB/张 |
| 数量限制 | 照片5张/人，附件10个/条目，缴费/退费证明各5张/笔，考勤1000条/次，选课500人/课 |
| 查询上限 | 考勤1000条，缴费5000条 |
| 数据完整性 | UNIQUE约束 + 唯一索引 + 外键 + 级联删除 + 事务提交 |

---

## 7. 性能设计

| 优化项 | 说明 |
|--------|------|
| WAL模式 | `PRAGMA journal_mode=WAL`，提升并发读写性能 |
| LIMIT上限 | 考勤1000条、缴费5000条、课时5000条，防全表扫描 |
| 学生分页 | 每页20条，LIMIT+OFFSET 分页 |
| 批量加载 | `list_students` 一次加载全部监护人和照片，避免N+1 |
| 显式列名 | BLOB表显式指定列名，排除data列 |
| 证明图片批量附加 | `_attach_photos()` 单次 IN 查询构建映射，列表接口零 N+1；图片本体按需加载 |
| 子查询统计 | 课程报读人数用子查询 |
| 课时聚合 | `_lesson_stats_rows()` 单次 SQL 三表 LEFT JOIN 聚合出全部课时指标，避免逐学生 N+1 查询 |
| 索引优化 | student_no、teacher_no 唯一索引 + attendance.student_id、payments.student_id、enrollments.course_id、enrollments.student_id、operation_logs.created_at 查询索引 |
| 先校验后保存 | 照片/附件先全部校验再批量入库 |
| UUID文件名 | 避免冲突：`{id}_{uuid12}{ext}` |
| Chart.js CDN | 图表库CDN引入，不占用本地存储 |

---

## 8. 部署方案

### 8.1 环境

| 项目 | 要求 |
|------|------|
| Python | 3.8+ |
| OS | Windows / Linux / macOS |
| 依赖 | flask>=3.0.0, openpyxl>=3.1.0 |
| 浏览器 | Chrome/Edge/Safari/Firefox（含移动端） |

### 8.2 启动

```bash
pip install -r requirements.txt
python app.py
```

启动时自动：`init_db()` → 建表+迁移 → `app.run(host="0.0.0.0", port=5000)`

### 8.3 配置

| 配置项 | 默认值 | 修改方式 |
|--------|--------|----------|
| 端口 | 5000 | 环境变量 `PORT` |
| 模板热重载 | 开启 | `TEMPLATES_AUTO_RELOAD=True` |
| Debug | 关闭 | `debug=False` |
| 数据库路径 | `students.db` | 代码常量 `DB_PATH` |

### 8.4 备份与恢复

**目录约定**：备份文件统一存放 `backups/data/` 专用目录，与代码备份 `backups/code_backup_*` 分开，避免互相干扰。旧版散落在 `backups/` 根目录的 `.db` 文件会在首次列备份时被 `_migrate_legacy_backups()` 自动迁入 `data/`。

| 能力 | 接口 | 说明 |
|------|------|------|
| 一键备份 | `POST /api/settings/backup` | 复制 `students.db` → `backups/data/students_YYYYMMDD_HHMMSS.db` |
| 列表 | `GET /api/settings/backups` | 扫描 `backups/data/*.db`，返回文件名/大小/修改时间（倒序）；手动拷入的有效库同样被识别 |
| 下载 | `GET .../download` | `send_file(as_attachment=True)` 下载到本地；前端用隐藏 `<a>` 触发，**不新开页面**（避免被 IDE 内置预览接管弹窗） |
| 恢复 | `POST .../restore` | 见下方流程 |
| 删除 | `DELETE ...` | 删除指定备份 |

**恢复流程**（防误操作三重保险）：

1. **有效性校验**：`sqlite3.connect(src)` 打开备份，检查 `sqlite_master` 中是否含 `students` 与 `users` 表；非系统数据库直接 400 拒绝，避免用任意文件覆盖导致数据全毁。
2. **自动留存快照**：恢复前先把当前 `students.db` 复制为 `students_恢复前自动备份_YYYYMMDD_HHMMSS.db`，误恢复可再从该文件还原。
3. **WAL 落盘**：恢复前执行 `PRAGMA wal_checkpoint(TRUNCATE)`，将 WAL 中未落盘的内容写回主文件，保证快照与恢复结果完整一致。
4. 覆盖主库文件 → 写入 `operation_logs`（action=恢复备份，detail 记录快照文件名）→ 提示用户刷新页面并重新登录。

> 日志写入失败不视为恢复失败（已用 `try/except sqlite3.Error` 隔离），避免日志表异常误导用户。

**前端渲染**：备份列表由 `loadBackups()` 动态拉取渲染，在**切换到「数据备份」选项卡时自动刷新**，备份/删除成功后局部刷新（不再整页 reload）。按钮采用事件委托 + 索引取文件名，规避文件名含空格/中文等特殊字符破坏 HTML 属性。

---

## 9. 文件清单

| 文件 | 说明 |
|------|------|
| `app.py` | 后端主程序（路由+API+DB+Excel导出+认证+日志+备份） |
| `students.db` | SQLite数据库（所有业务数据+BLOB，WAL模式） |
| `requirements.txt` | Python依赖 |
| `templates/login.html` | 登录页 |
| `templates/home.html` | 首页控制台 |
| `templates/index.html` | 学生档案列表 |
| `templates/resume.html` | 简历查看页 |
| `templates/resume_edit.html` | 简历编辑页 |
| `templates/attendance.html` | 考勤管理（考勤登记/补考勤/记录查询） |
| `templates/lessons.html` | 课时管理 |
| `templates/courses.html` | 课程管理 |
| `templates/teachers.html` | 教师管理 |
| `templates/payments.html` | 缴费管理 |
| `templates/refunds.html` | 退费管理 |
| `templates/subjects.html` | 科目管理 |
| `templates/import.html` | 数据导入 |
| `templates/reports.html` | 报表统计 |
| `templates/settings.html` | 系统设置 |
| `static/css/common.css` | 全局公共样式（CSS变量、导航栏、Toast、模态框） |
| `static/css/*.css` | 各页面样式文件（共 16 个，含 `lessons.css`） |
| `static/js/common.js` | 前端公共组件（Toast、Modal、fetch封装、表单校验） |
| `backups/data/` | 数据库备份专用目录（一键备份/下载/恢复） |
| `backups/code_backup_*` | 代码备份目录（按日期归档） |

---

## 10. 已知技术债务与优化方向

| 编号 | 问题 | 影响 | v2.0 状态 |
|------|------|------|----------|
| TD-01 | 无用户认证 | 任何内网用户可操作 | ✅ 已解决：Session认证 + 角色权限 |
| TD-02 | payments表残留 discount_percent/discount_amount 字段 | 冗余数据 | 新建库时移除，旧库保留兼容 |
| TD-03 | subjects表残留 price/lesson_count 字段 | 冗余数据 | 同上 |
| TD-04 | 教师与课程通过姓名文本匹配 | 改名后关联断裂 | ✅ 已解决：teacher_id 外键 |
| TD-05 | 单进程单线程 | 高并发时性能瓶颈 | 用gunicorn多worker部署 |
| TD-06 | 无日志系统 | 排错靠stderr文件 | ✅ 已解决：operation_logs + Python logging |
| TD-07 | 数据库无定期备份机制 | 数据丢失风险 | ✅ 已解决：一键备份功能 |
| TD-08 | 前端JS内联在HTML中 | 难以复用 | ✅ 已解决：common.js 公共组件 |
| TD-09 | 无自动化测试 | 回归靠人工 | 加pytest测试 |
| TD-10 | uploads目录仍保留但不再使用 | 冗余磁盘空间 | 确认迁移完成后清理 |

---

## 11. 版本记录

| 版本 | 日期 | 变更说明 |
|------|------|----------|
| v2.0 | 2026-09-14 | 用户认证、审计日志、退费、数据导入、系统设置、报表图表、考勤UPSERT、WAL、索引优化 |
| v3.0 | 2026-09-16 | UI 重构为 IDE 面板式科技感；「报读课程」更名「报读科目」并联动科目管理；简历页重做为朱砂宋体文书版式，支持导出 PDF；简历编辑页可直接编辑档案与监护人 |
| v3.1 | 2026-09-17 | 考勤模块重构：页面改为三选项卡（考勤登记按当天有课班级分班级卡展示、补考勤、记录查询）；attendance 表新增 start_time/end_time（上课时间段）与 source（来源：空=正常登记，makeup=补考勤）三列；考勤登记支持教师/上课时间/上课日期下拉（默认取课程管理数据）与临时学员；导出 Excel 新增「上课时间」列；移除月历视图、缺勤预警与月度统计接口 |
| v3.2 | 2026-10-01 | **新增课时管理模块**：`/lessons` 页面 + `/api/lessons`（含 filters/export）3 个接口；新增 `_lesson_stats_rows()` 统一课时核算（已购−已退−已消耗，已消耗=出勤+迟到考勤数），单次 SQL 三表 LEFT JOIN 聚合，被课时页/首页统计/学生列表三处复用；退费表新增 `lesson_count` 并校验累计退课时数不超过原缴费节数，新增退费导出；subjects 表新增 `year_price`/`lessons_per_year`。**备份升级**：备份目录收敛至 `backups/data/`（旧备份自动迁移），新增恢复接口（有效性校验 + 恢复前自动快照 + WAL checkpoint）与下载接口，备份列表改为动态拉取；新增索引 `enrollments.student_id`、`operation_logs.created_at`。 |
| v3.3 | 2026-10-01 | **缴费/退费证明图片**：新建 `payment_photos`/`refund_photos` 两张 BLOB 表（含外键索引）；新增上传/删除/查看下载共 6 个接口（API 总数 65→71）；新增通用函数 `_attach_photos()` 为列表批量附加图片元信息（单次 IN 查询，排除 data 列，零 N+1）；前端弹窗缩略图预览/标记删除/提交时同步上传，列表列点击隐藏 `<a>` 下载原图；删除缴费/退费/学生时级联清理图片，顺带修复删除学生不清退费记录的遗留缺陷。**全站移动端适配**（详见 4.4）：汉堡菜单抽屉、弹窗全屏（双类选择器覆盖各页样式）、表格/选项卡横滑、触控加大，全部收敛在 `common.css` 768px 断点 + `navbar.html` 公共组件，桌面端零影响。**修复**：用户管理「编辑」按钮 onclick 调用不存在的 `editUser`（实为 `openUserModal`）导致无法编辑。 |
| v3.4 | 2026-10-07 | **权限体系重构**：引入 principal（校长）唯一角色；全部权限收敛为 PERMISSION_MATRIX 权限点（新增 `attendance.batch`、`course.write`、`settings.logs`），前端模板全局 `can()` 裁剪入口、后端 `@permission_required` 拦截，补齐课程写接口（原无权限保护）与缴费/退费/科目等接口的权限装饰器；导航（数据报表/数据导入/系统设置）与首页模块卡片按权限显示，系统设置页标签页按角色裁剪（操作日志向教务/教师开放），并修复用户管理面板权限条件包裹缺陷导致的只读页面显示异常。**考勤**：新增 `POST /api/attendance/batch` 批量修改记录状态/备注（API 总数 71→72），记录查询页新增「修改查询信息状态」批量修改弹窗与「不限日期」快捷查询、默认范围放宽一年；只读角色考勤登记页课程信息与学员状态纯文本只读。**只读限制**：附件查看路由新增 `_viewer_download_denied()` 拦截（缴费/退费证明、学生照片、简历附件 `?download=1` 返回 403，预览不受限），简历编辑页路由加 `@resume_write_required`，简历页导出/编辑入口按权限渲染。**日志与账户联动**：删除用户级联删除其操作日志；修复登录日志在 Session 建立前写入导致的 user_id 缺失。**登录页**：移除默认账号提示文字。 |
