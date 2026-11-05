# 培训机构管理系统 — 产品需求文档（PRD）

| 项目 | 说明 |
|------|------|
| 文档版本 | v3.4 |
| 创建日期 | 2026-09-14 |
| 更新日期 | 2026-10-07 |
| 产品名称 | 培训机构管理系统 |
| 技术栈 | Python Flask + SQLite + HTML/CSS/JS + Chart.js（前后端不分离） |
| 部署方式 | 局域网内网部署，浏览器访问 `http://<服务器IP>:5000` |

---

## 1. 产品概述

### 1.1 产品定位

面向中小型教育培训机构的一体化管理工具，覆盖**用户认证、学生档案、教师管理、课程管理、考勤管理、课时管理、缴费管理、退费管理、科目管理、简历档案、报表统计、数据导入、系统设置**十三大核心业务模块。系统采用轻量级架构，无需安装客户端，局域网内任意设备（电脑/手机/平板）通过浏览器即可使用。

### 1.2 目标用户

- **系统管理员**：管理用户账号、系统设置、数据备份与恢复、查看操作日志
- **行政管理人员**：建档、缴费、退费、考勤录入、课时查询、数据导入
- **教师或教务负责人**：查看学生档案、课程排课、考勤记录、学员剩余课时
- **机构负责人**：查看报表统计、经营数据看板

### 1.3 核心价值

- **数据集中化**：所有业务数据（含照片、附件二进制文件）统一存入 SQLite 数据库，便于备份与迁移
- **安全可控**：用户认证 + 角色权限 + 操作审计日志，保障数据安全与可追溯
- **操作便捷**：浏览器即开即用，界面直观，Toast 通知 + 模态框交互，无需培训
- **功能完整**：从学生建档到缴费、退费、考勤、课时、报表、数据导入，覆盖培训机构日常运营全流程
- **课时可控**：自动核算学员「已购 - 已退 - 已消耗」剩余课时，课时用尽自动预警，支持按班级/教师筛选与导出
- **数据导出**：考勤、缴费、退费、课时均支持按条件筛选并导出格式化 Excel
- **数据安全**：数据库支持一键备份、下载到本地、一键恢复，恢复前自动留存当前数据快照
- **可视化报表**：报表页内置图表（柱状图/饼图/折线图），数据一目了然

---

## 2. 系统架构

### 2.1 技术架构

| 层级 | 技术选型 |
|------|----------|
| 后端框架 | Python Flask 3.0+ |
| 数据库 | SQLite（单文件 `students.db`，WAL 模式） |
| 前端 | 原生 HTML + CSS + JavaScript（无前端框架） |
| 图表 | Chart.js（CDN 引入，柱状图/饼图/折线图） |
| Excel 导出 | openpyxl 3.1+ |
| 文件存储 | 数据库 BLOB（照片与附件均存入数据库，不依赖磁盘文件） |
| 密码加密 | werkzeug.security（PBKDF2-SHA256） |
| 会话管理 | Flask Session（服务端签名 Cookie） |

### 2.2 数据库表结构

| 表名 | 说明 | 关键字段 |
|------|------|----------|
| `users` | 系统用户 | `username`(唯一), `password_hash`, `role`(admin/principal/staff/teacher/viewer), `display_name`, `status`(启用/禁用), `last_login` |
| `operation_logs` | 操作审计日志 | `user_id`, `username`, `action`(创建/修改/删除/登录/导出/备份), `module`, `target`, `detail`, `ip`, `created_at` |
| `students` | 学生档案 | `student_no`(唯一编号), `name`, `gender`, `birthdate`, `id_card`, `phone`, `school`, `grade`, `course`, `address`, `notes` |
| `guardians` | 监护人信息 | `student_id`(外键), `name`, `phone`, `wechat` |
| `photos` | 学生照片 | `student_id`, `filename`, `original_name`, `sort`(证件照排序), `data`(BLOB) |
| `resume_items` | 简历条目 | `student_id`, `type`(教育/培训/项目/证书/获奖/作品), `title`, `subtitle`, `date_text`, `description` |
| `attachments` | 简历附件 | `student_id`, `item_id`, `filename`, `original_name`, `size`, `data`(BLOB) |
| `attendance` | 考勤记录 | `student_id`, `date`, `status`(出勤/迟到/请假/缺勤), `course_name`, `lesson`, `teacher`, `start_time`, `end_time`, `source`(normal/makeup) |
| `courses` | 课程 | `name`, `category`, `teacher_id`(外键), `teacher`(冗余姓名), `schedule`, `hours`, `price`, `capacity`, `start_date`, `end_date`, `status` |
| `enrollments` | 选课记录 | `course_id`, `student_id` |
| `teachers` | 教师 | `teacher_no`(唯一编号), `name`, `gender`, `phone`, `subject`, `title`, `hire_date`, `status` |
| `payments` | 缴费记录 | `student_id`, `course`, `amount`, `unit_price`, `lesson_count`, `pay_date`, `method`, `remark`, `status`(已确认/待确认) |
| `payment_photos` | 缴费证明图片 | `payment_id`, `filename`, `original_name`, `size`, `data`(BLOB) |
| `refunds` | 退费记录 | `student_id`, `payment_id`, `amount`, `lesson_count`(退课时数), `refund_date`, `method`, `reason`, `remark` |
| `refund_photos` | 退费证明图片 | `refund_id`, `filename`, `original_name`, `size`, `data`(BLOB) |
| `subjects` | 科目 | `name`, `lesson_price`(课单价), `year_price`(年费), `lessons_per_year`(年课时数) |
| `system_settings` | 系统设置 | `key`, `value`, `updated_at` |
| `counters` | 编号计数器 | `name`, `value`(学生编号与教师工号自增序列) |

### 2.3 编号规则

| 编号类型 | 格式 | 规则 |
|----------|------|------|
| 学号 | `a00001` ~ `z99999` | 字母 a-z + 五位数字，自动递增，不可手动修改 |
| 教师工号 | `a0001` ~ `z9999` | 字母 a-z + 四位数字，自动递增 |

---

### 2.4 角色与权限点体系（v3.4 重构）

系统共 5 个角色：**admin（管理员）/ principal（校长）/ staff（教务）/ teacher（教师）/ viewer（只读）**。管理员与校长各全系统唯一。

权限采用**权限点**集中管理（`app.py` 的 `PERMISSION_MATRIX` 定义，前端模板 `can()` 与后端 `@permission_required` 装饰器共用同一份定义），页面按权限裁剪功能入口，无权限的功能不显示、接口同步拦截：

| 权限点 | 说明 | 允许角色 |
|--------|------|----------|
| `student.create` / `student.write` | 学生建档 / 修改删除档案 | admin、principal、staff、teacher（仅建档）；admin、principal、staff（修改删除） |
| `resume.write` | 编辑学生简历 | admin、principal、staff、teacher |
| `course.write` | 课程新建 / 编辑 / 删除 / 学员名单同步 | admin、principal、staff、teacher |
| `attendance.save` | 考勤登记 / 补考勤 | admin、principal、staff、teacher |
| `attendance.batch` | 记录查询页批量修改考勤状态/备注 | admin、principal、staff |
| `attendance.delete` | 删除考勤记录 | admin、principal、staff |
| `attendance.export` | 导出考勤 Excel | admin、principal、staff |
| `lesson.export` | 导出学员课时 Excel | admin、principal、staff |
| `payment.write` / `payment.delete` / `payment.export` | 缴费增改 / 删除 / 导出 | admin、principal、staff（删除与导出不含教师） |
| `refund.write` / `refund.delete` / `refund.export` | 退费增改 / 删除 / 导出 | admin、principal、staff（删除与导出不含教师） |
| `subject.write` / `subject.delete` | 科目增改 / 删除 | admin、principal、staff；删除仅 admin、principal |
| `teacher.create` / `teacher.write` | 教师新建 / 修改删除 | admin、principal、staff；修改删除仅 admin、principal |
| `report.view` | 查看数据报表 | admin、principal |
| `import.all` | 数据导入 | admin、principal、staff |
| `settings.manage` | 系统设置管理（用户/基本设置/备份） | admin、principal |
| `settings.logs` | 查看操作日志 | admin、principal、staff、teacher |

**只读角色（viewer）专项限制**：全部业务数据仅可查看——学生档案（无新建/编辑/删除按钮、无照片下载）、课程管理（无新建/编辑/删除入口）、考勤登记页课程信息（教师/上课时间/上课日期/本节课名称）与学员状态均为只读展示；缴费/退费证明图片、学生照片、简历附件**不可下载**（后端对 `?download=1` 附件请求返回 403，在线预览不受限）；简历页无「导出为 PDF」与「编辑简历」入口。

**导航与入口按权限显示**：
- 「数据报表」菜单与首页「报表统计」卡片：仅 `report.view`（管理员/校长）
- 「数据导入」菜单与首页卡片：仅 `import.all`（管理员/校长/教务）
- 「系统设置」菜单与首页卡片：所有登录角色可见，页内功能按权限裁剪（非管理员仅可修改密码）

#### 角色权限矩阵

| 模块/操作 | admin | principal | staff | teacher | viewer |
|-----------|:---:|:---:|:---:|:---:|:---:|
| 学生档案 - 新建 | ✓ | ✓ | ✓ | ✓ | ✗ |
| 学生档案 - 修改/删除 | ✓ | ✓ | ✓ | ✗ | ✗ |
| 学生档案 - 查看 | ✓ | ✓ | ✓ | ✓ | ✓ |
| 简历编辑 | ✓ | ✓ | ✓ | ✓ | ✗ |
| 课程管理 - 增删改 | ✓ | ✓ | ✓ | ✓ | ✗ |
| 课程管理 - 查看 | ✓ | ✓ | ✓ | ✓ | ✓ |
| 考勤管理 - 登记/补考勤 | ✓ | ✓ | ✓ | ✓ | ✗ |
| 考勤管理 - 批量修改记录 | ✓ | ✓ | ✓ | ✗ | ✗ |
| 考勤管理 - 删除记录 | ✓ | ✓ | ✓ | ✗ | ✗ |
| 考勤/课时/缴费/退费 - 导出 | ✓ | ✓ | ✓ | ✗ | ✗ |
| 缴费/退费/科目/教师 - 增删改 | ✓ | ✓ | ✓ | ✗/部分 | ✗ |
| 证明图片 - 在线查看 | ✓ | ✓ | ✓ | ✓ | ✓ |
| 证明图片 - 下载 | ✓ | ✓ | ✓ | ✓ | ✗ |
| 数据报表 | ✓ | ✓ | ✗ | ✗ | ✗ |
| 数据导入 | ✓ | ✓ | ✓ | ✗ | ✗ |
| 系统设置 - 管理 | ✓ | ✓ | ✗ | ✗ | ✗ |
| 操作日志 - 查看 | ✓ | ✓ | ✓ | ✓ | ✗ |
| 修改自己密码 | ✓ | ✓ | ✓ | ✓ | ✓ |

---

## 3. 功能模块详细需求

### 3.1 用户认证系统（`/login`）

**页面路由**：
- `GET /login` — 登录页
- `GET /logout` — 退出登录

**API 接口**：

| 方法 | 路径 | 说明 |
|------|------|------|
| POST | `/api/auth/login` | 用户登录（用户名 + 密码） |
| GET | `/api/auth/me` | 获取当前登录用户信息 |
| POST | `/api/auth/change-password` | 修改密码（首次登录强制改密码） |
| GET | `/api/users` | 用户列表（仅 admin / principal） |
| POST | `/api/users` | 新建用户（仅 admin / principal） |
| PUT | `/api/users/<id>` | 修改用户（仅 admin / principal；传 password 可重置密码） |
| DELETE | `/api/users/<id>` | 删除用户（仅 admin / principal，不可删除自己与唯一角色持有者；**删除时级联删除该用户全部操作日志**） |

**功能需求**：

#### 3.1.1 登录
- 用户名 + 密码登录，密码使用 PBKDF2-SHA256 加密存储
- 登录成功后建立 Session，并写入 `operation_logs`（action=登录，user_id 与用户名正确关联到该账户）
- 连续 5 次密码错误锁定该账号 15 分钟
- 系统初始化时自动创建默认管理员账号（admin / admin123），首次登录强制改密码
- 所有页面路由和 API 均需登录（除 `/login` 和 `/api/auth/login`）
- Session 有效期 8 小时，超时自动退出
- 登录页仅展示登录表单（v3.4 起移除登录按钮下方的默认账号提示文字）

#### 3.1.2 用户管理（仅 admin / principal）
- 用户字段：用户名、显示名、角色（admin/principal/staff/teacher/viewer）、状态（启用/禁用）
- 新建用户时设置初始密码；编辑弹窗内填写新密码可重置密码（留空不改）
- 支持修改角色、禁用/启用账号；管理员与校长全系统各唯一，不可通过编辑把自己降级或禁用
- 不可删除当前登录的自己、默认管理员账号与唯一角色持有者
- **删除用户时级联删除该用户的全部操作日志**（日志跟随账户）；「删除用户」这一操作本身记入删除者名下的日志，保证审计可追溯

---

### 3.2 首页 / 控制台（`/`）

**页面路由**：`GET /`

**功能描述**：
- 展示机构整体概况统计
- 显示最近建档的 6 名学生
- 提供各功能模块的导航入口

**展示数据项**：

| 统计项 | 说明 |
|--------|------|
| 学生总数 / 男 / 女 | 学生档案总量及性别分布 |
| 本月新生数 | 当月新建档学生数 |
| 照片总数 | 数据库中照片记录总数 |
| 附件总数 | 数据库中附件记录总数 |
| 考勤记录数 | 考勤记录总量 |
| 课程总数 | 开课课程总量 |
| 科目总数 | 科目档案总量 |
| 教师总数 | 教师档案总量 |
| 缴费总额 | 所有缴费记录金额合计 |
| 本月缴费 | 当月缴费金额合计 |
| 今日缴费 | 今日缴费金额合计 |
| 剩余课时合计 | 全部学员「已购-已退-已消耗」课时之和 |
| 课时预警人数 | 剩余课时 ≤ 0 的学员人数 |

---

### 3.2 学生档案管理（`/students`）

**页面路由**：`GET /students`
**API 接口**：

| 方法 | 路径 | 说明 |
|------|------|------|
| GET | `/api/students` | 学生列表（支持关键字搜索 + 分页 + 排序） |
| POST | `/api/students` | 新建学生档案 |
| PUT | `/api/students/<id>` | 修改学生档案 |
| DELETE | `/api/students/<id>` | 删除学生档案（级联删除） |
| GET | `/api/import/template` | 下载导入模板 Excel |
| POST | `/api/import/students` | 批量导入学生（Excel 上传） |

**功能需求**：

#### 3.2.1 学生列表与搜索
- 默认按建档时间倒序展示
- 支持分页：每页 20 条，支持翻页
- 搜索栏支持按以下字段模糊匹配：姓名、学号、电话、身份证号、课程、学校、监护人姓名/电话/微信
- 支持表头点击排序（姓名、建档日期、学号）
- 支持多条件筛选（性别、年级）
- 支持批量选择 + 批量删除
- 每条学生记录附带监护人和照片信息
- 每条学生附带课时汇总：`total_lessons`(已购)、`refunded_lessons`(已退)、`used_lessons`(已消耗)、`remaining_lessons`(剩余)

#### 3.2.2 建档（新增学生）
- 必填项：姓名（前端实时校验，输入即提示）
- 选填项：性别（默认男）、出生日期、身份证号（18位格式校验）、电话、学校、年级、课程、地址、备注
- 监护人信息：支持多条（姓名、电话、微信），全空的监护人自动忽略
- 学号自动生成，不可手动指定
- 身份证号格式校验：18位，末位可为 X
- 提交按钮防重复点击（提交后禁用 + loading 状态）

#### 3.2.3 修改档案
- 所有字段均可修改
- 监护人信息整体替换（先删后插）
- 修改时更新 `updated_at` 时间戳
- 修改操作写入 `operation_logs`

#### 3.2.4 删除档案
- 级联删除该学生的所有关联数据：照片、附件、简历条目、考勤、缴费、选课、监护人
- 删除操作写入 `operation_logs`
- 删除确认使用自定义模态框（非 `confirm()`），带红色危险提示

#### 3.2.5 照片管理

| 接口 | 说明 |
|------|------|
| `POST /api/students/<id>/photos` | 上传照片（最多 5 张/人，单张最大 5MB，支持 jpg/png/gif/webp） |
| `GET /api/students/<id>/photos` | 获取照片列表 |
| `DELETE /api/students/<id>/photos/<photo_id>` | 删除单张照片 |
| `POST /api/students/<id>/photos/<photo_id>/set-avatar` | 设为证件照（排在第一位） |
| `GET /uploads/photos/<filename>` | 查看照片（支持 `?download=1` 下载原始文件名） |

- 照片二进制数据存入数据库 `photos.data` BLOB 列
- 排序字段 `sort`：-1 表示证件照，0 表示普通照片

#### 3.2.6 批量导入学生
| 接口 | 说明 |
|------|------|
| `GET /api/import/template` | 下载导入模板 Excel |
| `POST /api/import/students` | 批量导入学生（Excel 上传） |

- 导入模板字段：姓名*、性别、出生日期、身份证号、电话、学校、年级、课程、地址、备注、监护人姓名、监护人电话、监护人微信、**已购课时**、**缴费金额(元)**、**缴费日期**
- 「已购课时」「缴费金额」「缴费日期」为选填，填写后导入时自动为该学生生成一条初始缴费记录（剩余课时即由此而来）
- 缴费日期须为 `YYYY-MM-DD` 格式，格式不合法时按导入当天处理
- 导入时逐行校验，跳过错误行并汇总报告
- 导入结果返回：成功数、失败数、失败原因列表
- 导入操作写入 `operation_logs`

---

### 3.3 学生简历档案（`/resume/<id>`）

**页面路由**：
- `GET /resume/<id>` — 简历查看页（可打印 / 保存为 PDF）
- `GET /resume/<id>/edit` — 简历编辑页（权限 `resume.write`，只读角色访问被拒绝）

**权限**：
- 「编辑简历」入口：仅管理员/校长/教务/教师（`resume.write`）
- 「导出为 PDF」按钮：除只读角色外均可使用
- 只读角色（viewer）：简历页仅可查看，工具栏不渲染

**功能需求**：

#### 3.3.1 简历展示
- 学生基本信息（姓名、学号、性别、年龄、学校、年级、课程）
- 身份证号脱敏显示（前6位 + 8个星号 + 后4位）
- 年龄自动计算：优先从出生日期计算，出生日期缺失时从身份证号推算
- 监护人信息
- 证件照展示
- 自我评价（对应学生档案的备注字段）
- 简历条目按 6 大类型分组展示
- 打印样式优化：`@media print` 控制分页，避免内容断裂

#### 3.3.2 简历条目管理

| 接口 | 说明 |
|------|------|
| `GET /api/students/<id>/resume-items` | 获取简历条目（按类型分组，含附件） |
| `POST /api/students/<id>/resume-items` | 新增条目 |
| `PUT /api/resume-items/<item_id>` | 修改条目 |
| `DELETE /api/resume-items/<item_id>` | 删除条目（附件一并删除） |

**简历条目类型**：

| 类型代码 | 中文名 |
|----------|--------|
| `education` | 教育经历 |
| `training` | 培训经历 |
| `project` | 项目经验 |
| `certificate` | 证书获取 |
| `award` | 竞赛获奖 |
| `work` | 作品展示 |

每条条目字段：标题（必填）、副标题、日期、描述

#### 3.3.3 附件管理

| 接口 | 说明 |
|------|------|
| `POST /api/resume-items/<item_id>/attachments` | 上传附件（每条目最多 10 个，单个最大 50MB，不限格式） |
| `DELETE /api/attachments/<attachment_id>` | 删除附件 |
| `GET /uploads/files/<filename>` | 下载附件（支持 `?download=1` 以原始文件名下载） |

- 附件二进制数据存入数据库 `attachments.data` BLOB 列
- 附件自动识别文件类型并显示图标（IMG/PDF/DOC/XLS/PPT/ZIP/VID/AUD/TXT/FILE）
- 图片类型附件支持在线预览

#### 3.3.4 自我评价
| 接口 | 说明 |
|------|------|
| `PUT /api/students/<id>/notes` | 修改自我评价（最多 500 字） |

---

### 3.4 考勤管理（`/attendance`）

**页面路由**：`GET /attendance`

**页面结构**：三个选项卡 — 考勤登记（默认）、补考勤、记录查询

**API 接口**：

| 方法 | 路径 | 说明 |
|------|------|------|
| GET | `/api/attendance?date=YYYY-MM-DD` | 获取某日考勤记录 |
| POST | `/api/attendance` | 批量保存考勤（同一学生同日仅一条，重复保存自动覆盖；含补考勤记录） |
| POST | `/api/attendance/batch` | 批量修改查询出的考勤记录的状态/备注（v3.4 新增，权限 `attendance.batch`） |
| DELETE | `/api/attendance/<record_id>` | 删除单条考勤记录 |
| GET | `/api/attendance/history` | 考勤记录查询（日期范围 + 状态多选 + 学生关键字 + source 来源筛选） |
| GET | `/api/attendance/export` | 导出考勤记录为 Excel |

**考勤状态**：

| 状态码 | 中文 | 是否计入总课时 |
|--------|------|---------------|
| `present` | 出勤 | 是 |
| `late` | 迟到 | 是 |
| `leave` | 请假 | 否 |
| `absent` | 缺勤 | 否 |

**考勤记录字段**：课程名称、本节课名称、上课老师、开始/结束时间（HH:MM，30 分钟档位下拉）、状态、备注、来源（`source`：空 = 考勤登记，`makeup` = 补考勤）

#### 3.4.1 考勤登记
- 选择考勤日期，自动列出当天有课的班级（判定依据：考勤日期在课程开课日期~结课日期范围内，且课程上课时间 schedule 含当天星期；无法解析星期的课程不展示）
- 每个班级一张班级卡，含：班级名、学员数、教师下拉（选项来自教师管理，默认为课程自带教师）、上课时间下拉（开始/结束，默认从课程 schedule 解析）、上课日期下拉（默认为考勤日期，可选排课周期内符合上课星期的其他日期）、上课人数
- 班级卡支持「本节课名称」输入，保存时写入本班每位学员的考勤记录
- 学员名单默认收起，点击「展开学员名单」查看；每行含学号、学员（头像+姓名）、考勤状态四选一按钮、备注
- 支持本班「全部出勤」快捷按钮
- 支持「添加临时学员」：搜索学生档案（已排除本班学员）点击添加，仅本次考勤有效，不写入班级名单
- 同一学员报多个班时，一天只保留一条考勤记录，状态跨班联动
- 「保存本班考勤」只保存该班学员（正式+临时）中已选择状态的记录
- 顶部汇总条：当日有课班级数、已登记/总人数、各状态人次、未登记人数
- 单次保存上限 1000 条
- **只读角色（viewer）**：班级卡的教师/上课时间/上课日期/本节课名称与学员状态、备注均以纯文本只读展示（下拉与输入框不渲染），无保存/全部出勤/临时学员按钮

#### 3.4.2 补考勤
- 为缺登的学生补录考勤：弹窗内先搜索学生（姓名/学号）选定，再选课程（选定后自动带入课程教师与上课时间段）
- 填写上课日期、本节课名称、出勤状态、备注，保存后记录标记 `source=makeup`
- 补考勤列表：按日期范围查询补考勤记录，展示日期/学号/姓名/课程/本节课/上课时间/老师/状态/备注，支持单条删除

#### 3.4.3 考勤记录查询
- 筛选条件：日期范围（起止日期，**默认最近一年**，可点「不限日期」一键查询全部）、状态多选（出勤/迟到/请假/缺勤，不选=全部）、学生关键字（姓名或学号）
- 查询结果含各状态统计与到课率（到课率 =（出勤+迟到）/记录总数）
- 汇总栏「到课率」右侧提供「修改查询信息状态」按钮（v3.4 新增，权限 `attendance.batch` = 管理员/校长/教务）：弹窗内勾选本次查询出的记录（默认全选，支持全选/反选/清空），统一修改为指定状态（不选=保持原状态）或修改/清空备注，保存后自动重新查询；修改状态会实时联动学员课时消耗（已消耗课时 = 出勤+迟到记录数）
- 查询上限 1000 条
- 支持单条删除（权限 `attendance.delete`）

#### 3.4.4 考勤导出 Excel
- 严格使用最近一次「查询」的条件导出，导出内容与页面查询结果逐条一致
- Excel 内容：标题行、筛选条件说明（含首位学员与家长信息）、明细表头（日期/学号/姓名/课程/本节课/上课时间/老师/状态/备注）、明细数据行、统计汇总、总课时（出勤+迟到）
- 文件名格式：`考勤记录_YYYYMMDD_HHMMSS.xlsx`（导出时刻）

---

### 3.5 课程管理（`/courses`）

**页面路由**：`GET /courses`

**API 接口**：

| 方法 | 路径 | 说明 |
|------|------|------|
| GET | `/api/courses` | 课程列表（支持关键字 + 状态筛选，含报读人数与学员名单） |
| POST | `/api/courses` | 新建课程（权限 `course.write`） |
| PUT | `/api/courses/<id>` | 修改课程（权限 `course.write`） |
| DELETE | `/api/courses/<id>` | 删除课程（级联删除选课记录，权限 `course.write`） |
| POST | `/api/courses/<id>/students` | 同步课程学员名单（全量覆盖，权限 `course.write`） |

**课程字段**：

| 字段 | 说明 | 必填 |
|------|------|------|
| name | 课程名称 | 是 |
| category | 课程类别 | 否 |
| teacher_id | 授课教师ID（外键关联 teachers.id） | 否 |
| teacher | 授课教师姓名（冗余字段，选择教师时自动填入） | 否 |
| schedule | 上课时间 | 否 |
| hours | 课时数 | 否 |
| price | 课程价格 | 否 |
| capacity | 容量上限 | 否 |
| start_date | 开课日期 | 否（默认创建当天） |
| end_date | 结课日期 | 否 |
| description | 课程描述 | 否 |
| status | 状态 | 默认"开课中"，可选"开课中"/"已结课" |

**校验规则**：
- 结课日期不能早于开课日期
- 开课日期留空时修改沿用原值
- 学员名单同步为全量覆盖模式，上限 500 人
- 授课教师通过下拉选择（从 teachers 表加载），选中后同时存储 `teacher_id` 和 `teacher`（姓名）
- 报读人数接近或超过 `capacity` 时前端显示橙色/红色预警
- 课程状态视觉区分：开课中=绿色标签，已结课=灰色标签
- 学员名单同步前显示二次确认弹窗，展示新增/移除的学员差异

---

### 3.6 教师管理（`/teachers`）

**页面路由**：`GET /teachers`

**API 接口**：

| 方法 | 路径 | 说明 |
|------|------|------|
| GET | `/api/teachers` | 教师列表（支持关键字 + 状态筛选，含授课门数） |
| POST | `/api/teachers` | 新建教师 |
| PUT | `/api/teachers/<id>` | 修改教师 |
| DELETE | `/api/teachers/<id>` | 删除教师 |

**教师字段**：

| 字段 | 说明 | 必填 |
|------|------|------|
| teacher_no | 教师工号 | 自动生成 |
| name | 姓名 | 是 |
| gender | 性别 | 默认男 |
| phone | 电话 | 否 |
| subject | 科目 | 否 |
| title | 职称 | 否 |
| hire_date | 入职日期 | 否 |
| status | 状态 | 默认"在职"，可选"在职"/"离职" |
| notes | 备注 | 否 |

- 关键字搜索支持：工号、姓名、科目、电话
- 授课门数：统计 `courses` 表中 `teacher_id` 匹配该教师 ID 的课程数

---

### 3.7 缴费管理（`/payments`）

**页面路由**：`GET /payments`

**API 接口**：

| 方法 | 路径 | 说明 |
|------|------|------|
| GET | `/api/payments` | 缴费记录列表（日期范围 + 学生关键字） |
| GET | `/api/payments/query` | 缴费查询（日期范围 + 学生关键字 + 科目 + 缴费方式） |
| GET | `/api/payments/export` | 导出缴费查询结果为 Excel |
| POST | `/api/payments` | 新建缴费记录 |
| PUT | `/api/payments/<id>` | 修改缴费记录 |
| DELETE | `/api/payments/<id>` | 删除缴费记录 |
| POST | `/api/payments/<id>/photos` | 上传缴费证明图片（最多 5 张/笔） |
| DELETE | `/api/payment-photos/<photo_id>` | 删除单张缴费证明 |
| GET | `/uploads/payment-photos/<filename>` | 查看缴费证明（`?download=1` 下载原图） |
| GET | `/api/subjects/lookup?name=xxx` | 按科目名称查询课单价（用于自动填入） |

**缴费字段**：

| 字段 | 说明 | 必填 |
|------|------|------|
| student_id | 缴费学生 | 是 |
| course | 科目 | 否 |
| amount | 缴费金额 | 是（必须 > 0） |
| unit_price | 科目单价 | 否 |
| lesson_count | 课程节数 | 否 |
| pay_date | 缴费日期 | 否（默认当天） |
| method | 缴费方式 | 默认"现金"，可选：现金/微信/支付宝/银行转账/其他 |
| status | 缴费状态 | 默认"已确认"，可选"已确认"/"待确认" |
| remark | 备注 | 否（最多 100 字） |

**缴费方式选项**：现金、微信、支付宝、银行转账、其他

#### 3.7.1 缴费自动计算
- 选择学生后自动加载该学生已报读课程列表
- 选择科目后自动从 `subjects` 表查询 `lesson_price` 填入 `unit_price`
- 输入 `unit_price` 和 `lesson_count` 后自动计算 `amount = unit_price × lesson_count`
- 自动计算可手动覆盖，用户修改 `amount` 后不再自动覆盖

#### 3.7.2 缴费查询
- 筛选条件：日期范围、学生关键字（姓名或学号）、科目（精确匹配）、缴费方式（精确匹配）
- 查询上限 5000 条
- 表格显示列：缴费日期、学号、学生姓名、科目、课程节数、缴费金额、缴费方式、备注
- 缴费列表底部显示当前页金额小计行
- 缴费方式列用彩色标签（Chip）区分不同方式

#### 3.7.3 缴费导出 Excel
- 按查询条件导出
- Excel 内容：标题行、筛选条件说明、明细表头（缴费日期/学号/学生姓名/科目/科目单价/课程节数/缴费金额/缴费方式/备注）、明细数据行、合计行（总金额 + 笔数）
- 文件名格式：`缴费查询_YYYYMMDD_HHMMSS.xlsx`

#### 3.7.4 学生选择器
- 缴费时学生选择改为可搜索下拉选择器
- 输入姓名/学号自动匹配

#### 3.7.5 缴费证明图片
- 新建/编辑缴费弹窗最后提供「缴费证明」上传项：**选填**，每笔最多 5 张，支持 jpg/png/gif/webp，单张不超过 5MB
- 选择后即时显示缩略图预览，每张可点 × 移除；编辑时回显已有图片
- 编辑时移除已有图片为「标记删除」，提交时才真正删除（取消编辑不误删）
- 提交流程：先保存缴费记录，再用返回的记录 id 自动上传图片（新建场景）
- 图片二进制存入数据库 `payment_photos.data` BLOB 列，不依赖磁盘
- 缴费管理列表与记录查询列表均增加「缴费证明」列：显示缩略图（多张并排 + 数量角标），**点击缩略图即下载原图**（隐藏 `<a>` 触发，不弹新页面）
- **只读角色（viewer）可在线查看缩略图，但点击不触发下载**，后端对 `?download=1` 的附件请求返回 403（v3.4）
- 删除缴费记录、删除学生档案时级联删除其全部证明图片

---

### 3.8 科目管理（`/subjects`）

**页面路由**：`GET /subjects`

**API 接口**：

| 方法 | 路径 | 说明 |
|------|------|------|
| GET | `/api/subjects` | 科目列表 + 统计（科目数 / 平均课单价） |
| POST | `/api/subjects` | 新增科目 |
| PUT | `/api/subjects/<id>` | 修改科目 |
| DELETE | `/api/subjects/<id>` | 删除科目 |

**科目字段**：

| 字段 | 说明 | 必填 |
|------|------|------|
| name | 科目名称 | 是（不超过 50 字） |
| lesson_price | 每节课课单价 | 是（不小于 0 的数字） |

- 科目档案仅维护科目名称与课单价，用于缴费时引用
- 统计信息：科目总数、平均课单价

---

### 3.9 退费管理（`/refunds`）

**页面路由**：`GET /refunds`

**API 接口**：

| 方法 | 路径 | 说明 |
|------|------|------|
| GET | `/api/refunds` | 退费记录列表（日期范围 + 学生关键字） |
| GET | `/api/refunds/export` | 导出退费记录为 Excel |
| POST | `/api/refunds` | 新建退费记录（关联原缴费记录） |
| PUT | `/api/refunds/<id>` | 修改退费记录 |
| DELETE | `/api/refunds/<id>` | 删除退费记录 |
| POST | `/api/refunds/<id>/photos` | 上传退费证明图片（最多 5 张/笔） |
| DELETE | `/api/refund-photos/<photo_id>` | 删除单张退费证明 |
| GET | `/uploads/refund-photos/<filename>` | 查看退费证明（`?download=1` 下载原图） |

**退费字段**：

| 字段 | 说明 | 必填 |
|------|------|------|
| student_id | 退费学生 | 是 |
| payment_id | 关联原缴费记录 | 否 |
| amount | 退费金额 | 是（必须 > 0） |
| lesson_count | 退课时数 | 否（默认 0，须 ≥ 0） |
| refund_date | 退费日期 | 否（默认当天） |
| method | 退费方式 | 默认"原路退回"，可选：原路退回/现金/微信/支付宝/银行转账/其他 |
| reason | 退费原因 | 是 |
| remark | 备注 | 否 |

**功能需求**：
- 选择学生后自动展示该学生的缴费记录列表，可选择关联的原缴费记录
- 退费金额不可超过原缴费金额（校验同一缴费记录的累计退费金额，提示已退/本次/原缴费三项明细）
- 退课时数不可超过原缴费记录的课程节数（校验同一缴费记录的累计退课时数）
- 退费后该学员剩余课时相应减少，课时管理页与首页统计同步更新
- 报表统计中退费金额从缴费总额中扣减
- 支持按查询条件导出退费记录 Excel

#### 3.9.1 退费证明图片
- 与缴费证明同套交互：新建/编辑退费弹窗提供「退费证明」上传项（**选填**，每笔最多 5 张，jpg/png/gif/webp，单张 5MB 内）
- 缩略图预览、可移除；编辑回显；标记删除提交时生效；保存后自动上传（新建用返回的记录 id）
- 图片二进制存入 `refund_photos.data` BLOB 列
- 退费登记（最近列表）与记录查询列表均增加「退费证明」列：缩略图并排显示，**点击即下载原图**
- **只读角色（viewer）可在线查看缩略图，但不可下载**（后端拦截 `?download=1`，v3.4）
- 删除退费记录、删除学生档案时级联删除其全部证明图片（修复了旧版删除学生不清退费记录的遗留问题）

---

### 3.10 课时管理（`/lessons`）

**页面路由**：`GET /lessons`

**API 接口**：

| 方法 | 路径 | 说明 |
|------|------|------|
| GET | `/api/lessons/filters` | 筛选选项：班级（课程）列表 + 授课教师列表 |
| GET | `/api/lessons` | 学员课时列表（关键字 + 班级 + 教师 + 预警筛选，含汇总统计） |
| GET | `/api/lessons/export` | 导出学员课时为 Excel |

**功能定位**：以学员为维度汇总课时资产，回答「谁还有多少课没上、谁该续费了」。

#### 3.10.1 课时核算口径（全系统统一）

| 指标 | 计算方式 |
|------|----------|
| 已购课时 | `payments` 表该学生 `lesson_count` 之和 |
| 已退课时 | `refunds` 表该学生 `lesson_count` 之和 |
| 已消耗课时 | `attendance` 表该学生状态为**出勤 + 迟到**的记录数（请假/缺勤不扣课时） |
| 剩余课时 | 已购 − 已退 − 已消耗（可为负，表示课时超上） |
| 使用率 | 已消耗 ÷ (已购 − 已退) × 100%（分母 ≤ 0 时为 0） |
| 课时预警 | 剩余课时 ≤ 0（含未购课但已上课的情况） |

> 该口径同时用于：课时管理页、首页「剩余课时合计 / 课时预警人数」、学生档案列表的课时字段。

#### 3.10.2 查询与筛选
- 关键字：姓名 或 学号（模糊匹配）
- 班级：下拉选择课程，仅显示该班报读学员（来源 `enrollments`）
- 授课教师：下拉选择教师，显示其名下所有班级的学员（兼容 `teacher_id` 外键与 `teacher` 姓名两种关联）
- 「只看课时预警」勾选框：仅显示剩余课时 ≤ 0 的学员
- 结果按剩余课时升序排列（课时最少的排最前，便于优先跟进），上限 5000 条

#### 3.10.3 统计卡片
| 指标 | 说明 |
|------|------|
| 剩余课时合计 | 当前筛选结果剩余课时之和（核心指标，首位展示） |
| 课时预警人数 | 剩余课时 ≤ 0 的学员数 |
| 学员人数 | 当前筛选结果学员数 |
| 已购课时合计 | 已购课时之和 |
| 已消耗课时 | 已消耗课时之和 |

#### 3.10.4 列表字段
学号、姓名、剩余课时、使用率、状态（正常/预警标签）、已购课时、已退课时、已消耗课时、所属班级、授课教师、性别、年级、报读科目

#### 3.10.5 导出 Excel
- 严格使用当前筛选条件导出
- 文件名格式：`学员课时表_YYYYMMDD_HHMMSS.xlsx`

---

### 3.11 报表统计（`/reports`）

**页面路由**：`GET /reports`

**功能描述**：经营数据看板，一页展示全机构核心运营指标，内置图表可视化。

**展示数据项**：

| 分类 | 指标 | 说明 |
|------|------|------|
| 学生 | 总数 / 男 / 女 / 本月新增 | 性别分布与月度增长 |
| 学生 | 课程分布 | 按学生报读课程分组统计（Top 8，水平柱状图） |
| 课程 | 总数 / 开课中 | 课程总量与活跃数 |
| 教师 | 总数 / 在职 | 教师总量与在职数 |
| 缴费 | 总笔数 / 总金额 / 本月金额 / 今日金额 | 缴费规模与趋势 |
| 缴费 | 按科目统计 | 各科目缴费笔数与金额（Top 5，柱状图） |
| 缴费 | 按方式统计 | 各缴费方式笔数与金额（饼图） |
| 缴费 | 近6月趋势 | 按月统计缴费金额（折线图） |
| 退费 | 总笔数 / 总金额 | 退费汇总 |
| 考勤 | 总记录数 / 出勤率 | 全量考勤汇总 |
| 考勤 | 近 30 天出勤率 | 近期考勤趋势 |
| 考勤 | 各状态人次 | 出勤/迟到/请假/缺勤（环形图） |

**图表可视化**：
- 课程分布：水平柱状图
- 缴费按科目：柱状图
- 缴费按方式：饼图
- 缴费趋势：折线图
- 考勤状态分布：环形图

**时间范围**：
- 支持自定义日期范围筛选
- 快捷按钮：本月、近30天、近3个月、全部

---

### 3.12 系统设置（`/settings`，所有角色可进入，页内功能按权限裁剪）

**页面路由**：`GET /settings`

**页面结构**：选项卡按角色显示 —
- 管理员 / 校长：用户管理、基本设置、数据备份、操作日志、修改密码（默认用户管理）
- 教务 / 教师：操作日志、修改密码（默认操作日志）
- 只读：修改密码

**API 接口**：

| 方法 | 路径 | 说明 |
|------|------|------|
| GET/PUT | `/api/settings` | 读写机构基本设置（名称/电话/地址，仅管理员/校长） |
| POST | `/api/settings/backup` | 立即备份数据库（仅管理员/校长） |
| GET | `/api/settings/backups` | 备份文件列表（仅管理员/校长） |
| GET | `/api/settings/backups/<filename>/download` | 下载备份文件到本地（仅管理员/校长） |
| POST | `/api/settings/backups/<filename>/restore` | 用指定备份恢复数据库（仅管理员/校长） |
| DELETE | `/api/settings/backups/<filename>` | 删除备份文件（仅管理员/校长） |
| GET | `/api/settings/operation-logs` | 操作日志查询（权限 `settings.logs`：管理员/校长/教务/教师） |

#### 3.12.1 数据备份
- 点击「立即备份」将当前 `students.db` 复制为 `students_YYYYMMDD_HHMMSS.db`
- **备份存放位置**：服务器 `backups/data/` 专用目录（与代码备份 `code_backup_*` 目录分开，互不干扰）
- 首次访问时自动把旧版散落在 `backups/` 根目录的 `.db` 备份迁移进 `backups/data/`
- 备份列表展示文件名、创建时间、大小，按时间倒序
- 手动拷入 `backups/data/` 的有效 `.db` 文件也会被自动识别并显示在列表中
- 备份列表在**每次切换到「数据备份」选项卡时自动刷新**，备份/删除后局部刷新无需整页重载

#### 3.12.2 备份下载与恢复
- **下载**：点击「下载」以隐藏 `<a>` 标签触发浏览器下载，**不新开页面/弹窗**，保存为原文件名
- **恢复**：点击「恢复」弹出二次确认，确认后执行：
  1. 校验备份文件为有效 SQLite 库且含 `students`、`users` 表（非法文件直接拒绝并提示）
  2. 恢复前自动把当前数据库另存为 `students_恢复前自动备份_YYYYMMDD_HHMMSS.db`（可用于撤销本次恢复）
  3. 执行 `PRAGMA wal_checkpoint(TRUNCATE)` 将 WAL 内容落盘后再覆盖主库文件
  4. 成功后提示用户刷新页面并重新登录
- 恢复与删除操作均写入 `operation_logs`（恢复记录附带恢复前快照文件名）

#### 3.12.3 操作日志
- 权限 `settings.logs`：管理员/校长/教务/教师可查（v3.4 起向教务与教师开放）
- 教务/教师进入系统设置默认展示操作日志标签页
- 筛选条件：日期范围、用户名、操作类型、模块
- 展示：时间、用户、操作、模块、详情
- 默认不加载，点击「查询」后展示
- **日志跟随账户**：删除用户账户时，该用户的全部操作日志同步删除；「删除用户」操作本身记录在删除者名下

---

## 4. 非功能需求

### 4.1 数据存储
- 所有数据（含照片、附件二进制文件）统一存入 SQLite 数据库 `students.db`
- 数据库启用 WAL 模式，提升并发读性能
- 数据库单文件，便于整体备份与迁移
- 数据库表结构支持向后兼容（`init_db()` 自动检测并添加新列、迁移旧数据）

### 4.2 数据安全
- 用户认证：所有操作需登录，Session 签名 Cookie
- 密码加密：PBKDF2-SHA256 加密存储
- 身份证号在简历页脱敏显示（前6后4，中间星号）
- 操作审计日志：关键操作（创建/修改/删除/登录/导出/备份/恢复）记录操作人、时间、IP；**删除用户时级联删除其名下日志**
- **只读角色下载限制**：缴费/退费证明、学生照片、简历附件对只读角色禁止附件下载（`?download=1` 返回 403），在线预览不受限
- 账号锁定：连续 5 次密码错误锁定 15 分钟
- **备份可恢复**：备份文件专存 `backups/data/` 目录，支持一键恢复与下载到本地
- **恢复安全网**：恢复前自动把当前数据库另存为快照文件（`students_恢复前自动备份_YYYYMMDD_HHMMSS.db`），误恢复可再还原
- **恢复校验**：恢复前校验备份文件为有效 SQLite 库且含 `students`/`users` 表，非法文件直接拒绝，避免数据全毁
- **恢复一致性**：恢复前执行 `PRAGMA wal_checkpoint(TRUNCATE)` 确保 WAL 内容落盘
- 删除操作前端有确认提示（自定义模态框）
- SQL 注入防护：全部参数化查询
- XSS 防护：前端 `escapeHtml()` 转义

### 4.3 性能
- 查询接口返回上限：考勤 1000 条、缴费 5000 条、课时 5000 条、选课名单 500 人
- 学生列表分页：每页 20 条
- 照片限制：每学生最多 5 张，单张最大 5MB
- 附件限制：每简历条目最多 10 个，单个最大 50MB
- 证明图片限制：每笔缴费/退费最多 5 张，单张最大 5MB（jpg/png/gif/webp）
- 数据库索引优化：attendance.student_id、payments.student_id、enrollments.course_id 等关键查询字段建索引

### 4.4 兼容性
- 浏览器兼容：Chrome / Edge / Safari / Firefox（支持移动端浏览器）
- 操作系统兼容：服务端支持 Windows / Linux / macOS
- 模板文件修改后自动生效（`TEMPLATES_AUTO_RELOAD`），无需重启

### 4.5 移动端适配（v3.3 起）
全站响应式：**电脑端（>768px）体验保持不变**，手机端（≤768px）自动切换移动体验，改动全部收敛在公共组件中：

- **导航**：汉堡菜单（☰）按钮，点击展开深色抽屉（4 列网格链接、当前页高亮），点链接/遮罩自动收起；桌面端仍为水平导航
- **弹窗**：新建/编辑弹窗全屏铺开，表单单列、底部按钮拉通，适配 iPhone 底部安全区
- **表格**：数据表格容器横向惯性滚动（`-webkit-overflow-scrolling`），不挤压折行
- **选项卡**：横向滑动，禁止压缩成竖排文字
- **触控**：小按钮、分页页码加大至手指友好尺寸；Toast 顶部通栏
- 登录页、首页模块卡片、统计卡片此前已适配，保持不变

### 4.5 部署
- 启动命令：`python app.py`
- 默认端口：5000（可通过环境变量 `PORT` 修改）
- 局域网访问：需在防火墙中放行对应端口
- 启动时自动初始化数据库表结构与数据迁移

---

## 5. 页面导航结构

```
登录（/login）
首页 / 控制台（/）
├── 学生档案（/students）
│   ├── 简历查看（/resume/<id>）
│   └── 简历编辑（/resume/<id>/edit）
├── 考勤管理（/attendance）
├── 课时管理（/lessons）
├── 课程管理（/courses）
├── 教师管理（/teachers）
├── 缴费管理（/payments）
├── 退费管理（/refunds）
├── 科目管理（/subjects）
├── 数据导入（/import）
├── 报表统计（/reports）
└── 系统设置（/settings）
```

---

## 6. API 接口汇总

### 6.1 用户认证

| 方法 | 路径 | 功能 |
|------|------|------|
| POST | `/api/auth/login` | 用户登录 |
| GET | `/api/auth/me` | 获取当前用户信息 |
| POST | `/api/auth/change-password` | 修改密码 |
| GET | `/api/users` | 用户列表（仅 admin/principal） |
| POST | `/api/users` | 新建用户（仅 admin/principal） |
| PUT | `/api/users/<id>` | 修改用户（仅 admin/principal） |
| DELETE | `/api/users/<id>` | 删除用户（仅 admin/principal，级联删除其操作日志） |

### 6.2 学生档案

| 方法 | 路径 | 功能 |
|------|------|------|
| GET | `/api/students` | 学生列表（关键字搜索 + 分页 + 排序） |
| POST | `/api/students` | 新建学生档案 |
| PUT | `/api/students/<id>` | 修改学生档案 |
| DELETE | `/api/students/<id>` | 删除学生档案 |
| POST | `/api/students/<id>/photos` | 上传照片 |
| GET | `/api/students/<id>/photos` | 获取照片列表 |
| DELETE | `/api/students/<id>/photos/<photo_id>` | 删除照片 |
| POST | `/api/students/<id>/photos/<photo_id>/set-avatar` | 设为证件照 |
| PUT | `/api/students/<id>/notes` | 修改自我评价 |
| GET | `/api/import/template` | 下载导入模板 |
| POST | `/api/import/students` | 批量导入学生 |

### 6.2 简历条目与附件

| 方法 | 路径 | 功能 |
|------|------|------|
| GET | `/api/students/<id>/resume-items` | 获取简历条目 |
| POST | `/api/students/<id>/resume-items` | 新增条目 |
| PUT | `/api/resume-items/<id>` | 修改条目 |
| DELETE | `/api/resume-items/<id>` | 删除条目 |
| POST | `/api/resume-items/<id>/attachments` | 上传附件 |
| DELETE | `/api/attachments/<id>` | 删除附件 |

### 6.3 文件访问

| 方法 | 路径 | 功能 |
|------|------|------|
| GET | `/uploads/photos/<filename>` | 查看照片 |
| GET | `/uploads/files/<filename>` | 下载附件 |

### 6.4 考勤管理

| 方法 | 路径 | 功能 |
|------|------|------|
| GET | `/api/attendance?date=YYYY-MM-DD` | 获取某日考勤 |
| POST | `/api/attendance` | 批量保存考勤（含补考勤） |
| POST | `/api/attendance/batch` | 批量修改考勤记录状态/备注（v3.4 新增） |
| DELETE | `/api/attendance/<id>` | 删除考勤记录 |
| GET | `/api/attendance/history` | 考勤记录查询（`?source=makeup` 查补考勤） |
| GET | `/api/attendance/export` | 导出考勤 Excel |

### 6.5 课程管理

| 方法 | 路径 | 功能 |
|------|------|------|
| GET | `/api/courses` | 课程列表 |
| POST | `/api/courses` | 新建课程 |
| PUT | `/api/courses/<id>` | 修改课程 |
| DELETE | `/api/courses/<id>` | 删除课程 |
| POST | `/api/courses/<id>/students` | 同步学员名单 |

### 6.6 教师管理

| 方法 | 路径 | 功能 |
|------|------|------|
| GET | `/api/teachers` | 教师列表 |
| POST | `/api/teachers` | 新建教师 |
| PUT | `/api/teachers/<id>` | 修改教师 |
| DELETE | `/api/teachers/<id>` | 删除教师 |

### 6.7 缴费管理

| 方法 | 路径 | 功能 |
|------|------|------|
| GET | `/api/payments` | 缴费记录列表 |
| GET | `/api/payments/query` | 缴费查询 |
| GET | `/api/payments/export` | 导出缴费 Excel |
| POST | `/api/payments` | 新建缴费记录 |
| PUT | `/api/payments/<id>` | 修改缴费记录 |
| DELETE | `/api/payments/<id>` | 删除缴费记录 |
| POST | `/api/payments/<id>/photos` | 上传缴费证明 |
| DELETE | `/api/payment-photos/<id>` | 删除缴费证明 |
| GET | `/uploads/payment-photos/<filename>` | 查看/下载缴费证明 |
| GET | `/api/subjects/lookup?name=xxx` | 查询科目课单价 |

### 6.8 退费管理

| 方法 | 路径 | 功能 |
|------|------|------|
| GET | `/api/refunds` | 退费记录列表 |
| GET | `/api/refunds/export` | 导出退费 Excel |
| POST | `/api/refunds` | 新建退费记录 |
| PUT | `/api/refunds/<id>` | 修改退费记录 |
| DELETE | `/api/refunds/<id>` | 删除退费记录 |
| POST | `/api/refunds/<id>/photos` | 上传退费证明 |
| DELETE | `/api/refund-photos/<id>` | 删除退费证明 |
| GET | `/uploads/refund-photos/<filename>` | 查看/下载退费证明 |

### 6.9 科目管理

| 方法 | 路径 | 功能 |
|------|------|------|
| GET | `/api/subjects` | 科目列表 + 统计 |
| POST | `/api/subjects` | 新增科目 |
| PUT | `/api/subjects/<id>` | 修改科目 |
| DELETE | `/api/subjects/<id>` | 删除科目 |

### 6.10 课时管理

| 方法 | 路径 | 功能 |
|------|------|------|
| GET | `/api/lessons/filters` | 筛选选项（班级 + 授课教师） |
| GET | `/api/lessons` | 学员课时列表 + 汇总统计 |
| GET | `/api/lessons/export` | 导出学员课时 Excel |

### 6.11 系统设置

| 方法 | 路径 | 功能 |
|------|------|------|
| GET | `/api/settings` | 获取系统设置 |
| PUT | `/api/settings` | 修改系统设置 |
| POST | `/api/settings/backup` | 一键备份数据库 |
| GET | `/api/settings/backups` | 备份文件列表（读取 `backups/data/`） |
| GET | `/api/settings/backups/<filename>/download` | 下载备份文件到本地 |
| POST | `/api/settings/backups/<filename>/restore` | 恢复备份（覆盖当前数据库） |
| DELETE | `/api/settings/backups/<filename>` | 删除备份文件 |
| GET | `/api/settings/operation-logs` | 操作日志查询 |

---

## 7. 文件清单

| 文件 | 说明 |
|------|------|
| `app.py` | 后端主程序（路由、API、数据库操作、Excel 导出、认证、日志） |
| `students.db` | SQLite 数据库（所有业务数据，WAL 模式） |
| `requirements.txt` | Python 依赖（flask, openpyxl） |
| `templates/login.html` | 登录页 |
| `templates/home.html` | 首页 / 控制台 |
| `templates/index.html` | 学生档案列表 |
| `templates/resume.html` | 简历查看页 |
| `templates/resume_edit.html` | 简历编辑页 |
| `templates/attendance.html` | 考勤管理 |
| `templates/lessons.html` | 课时管理 |
| `templates/courses.html` | 课程管理 |
| `templates/teachers.html` | 教师管理 |
| `templates/payments.html` | 缴费管理 |
| `templates/refunds.html` | 退费管理 |
| `templates/subjects.html` | 科目管理 |
| `templates/import.html` | 数据导入 |
| `templates/reports.html` | 报表统计 |
| `templates/settings.html` | 系统设置 |
| `static/css/common.css` | 全局公共样式（CSS 变量、导航栏、Toast、模态框） |
| `static/css/*.css` | 各页面样式文件（共 16 个，含 `lessons.css`） |
| `static/js/common.js` | 前端公共组件（fetch 封装、Toast、Modal、表单校验） |
| `backups/data/` | 数据库备份专用目录（与代码备份分开） |
| `backups/code_backup_*` | 代码备份目录（按日期归档 `app.py`/`templates`/`static`） |

---

## 8. UI 设计规范

> v3.0（2026-09-16）起生效。设计方向：**IDE 面板式科技感**——视觉语言取自代码编辑器（深色侧栏 + 浅色编辑区 + 终端青 + 等宽字体），克制、专业，服务于编程培训机构的品牌气质。

### 8.1 设计 Token（`static/css/common.css` 的 `:root` 为唯一来源）

| 类别 | 变量 | 值 | 用途 |
|------|------|-----|------|
| 主色 | `--primary` | `#0891b2` | 信号青：主按钮、链接、聚焦态、导航 active 底条 |
| 主色-深 | `--primary-dark` | `#0e7490` | 主色 hover |
| 主色-浅底 | `--primary-light` | `#e6f7fa` | 浅青底色（hover 底、标签底） |
| 主色-浅边 | `--primary-border` | `#a5d8de` | 浅青边框 |
| 墨色 | `--ink` / `--ink-2` | `#0f172a` / `#1e293b` | 导航栏、KPI 卡、登录页背景 |
| 语义色 | `--success` / `--danger` / `--warning` / `--info` | `#16a34a` / `#dc2626` / `#d97706` / `#0891b2` | 状态表达 |
| 文本 | `--text` / `--text-secondary` / `--text-muted` | `#1f2937` / `#6b7280` / `#9ca3af` | 三级文字 |
| 边框/背景 | `--border` / `--bg` / `--bg-white` | `#e5e7eb` / `#f6f7f9` / `#ffffff` | hairline 与底色 |
| 圆角 | `--radius` / `--radius-sm` | `6px` / `4px` | 小圆角，告别胶囊 |
| 阴影 | `--shadow` / `--shadow-sm` | 单层浅阴影 | 阴影退居次要，边框为主 |
| 等宽字体 | `--mono` | Cascadia Code → JetBrains Mono → Consolas 回退 | 学号/金额/编号/统计数字/状态标签 |

### 8.2 设计原则

1. **导航栏**：深藏青扁平色带（无渐变），品牌区代码符号 `</>` logo（青色等宽），当前页用青色底条标识
2. **数据表达**：学号、金额、统计数字、状态标签统一等宽字体——"代码感"来自内容排版而非装饰
3. **形状语言**：圆角收敛至 4–8px；状态标签为方角小牌；hairline 边框为主、阴影大幅减弱
4. **色彩纪律**：全站无渐变装饰；主色仅信号青一族，语义色各司其职
5. **动效**：仅保留 hover/focus 反馈，无入场动画
6. **登录页**：深墨背景 + 青色微光光斑 + 半透明深色面板，青色聚焦态

### 8.3 例外

- 学生简历页（`resume.html` / `resume.css`）为交付家长的**打印文书**，采用独立的朱砂宋体文书版式（A4 排版、可导出 PDF），不属于后台 UI 体系，刻意保留独立风格。

### 8.4 适配说明

- 全站 14 个页面 CSS 统一引用 `common.css` 的 `:root` token，业务 CSS 不再重复定义变量
- 后续换肤只需修改 `common.css` 一处变量值即可全局生效

---

## 9. 版本记录

| 版本 | 日期 | 变更说明 |
|------|------|----------|
| v1.0 | 2026-09-14 | 初始版本：八大功能模块全部完成，数据存储全部入库（BLOB），Excel 导出功能完善 |
| v2.0 | 2026-09-14 | 新增用户认证系统、操作审计日志、退费管理、数据导入、系统设置、报表图表可视化、学生分页、缴费自动计算、考勤批量录入与月历视图、教师-课程外键关联、SQLite WAL 模式、数据库索引优化、前端公共组件（Toast/Modal/表单校验）、CSS 全局变量体系、统一导航栏 |
| v3.0 | 2026-09-16 | UI 全面重构为「IDE 面板式科技感」（详见第 8 章）：导航去渐变改深藏青、主色信号青、等宽字体数据展示、圆角与阴影收敛、全站去渐变、登录页深色科技面板；「报读课程」更名「报读科目」并与科目管理联动（学生档案、课程类别下拉数据均来自科目管理）；简历页重做为朱砂宋体文书版式并支持导出 PDF；简历编辑页可直接编辑档案信息与监护人 |
| v3.1 | 2026-09-17 | 考勤管理重构：页面改为三选项卡（考勤登记/补考勤/记录查询）；考勤登记按「当天有课班级」分班级卡展示（依据课程开结课日期与上课星期判定），支持教师/上课时间/上课日期下拉（默认取课程管理数据）、本节课名称整班写入、临时学员、班级卡「全部出勤」与按班保存；新增补考勤功能（source=makeup 标记，弹窗搜索学生补录）；考勤记录新增上课时间段（start_time/end_time）；导出 Excel 新增「上课时间」列、文件名改为按查询截止日期命名；移除月历视图、缺勤预警与月度统计接口 |
| v3.2 | 2026-10-01 | **新增课时管理模块**（第 13 个业务模块）：新增 `/lessons` 页面与 `/api/lessons` 系列 3 个接口，按「已购−已退−已消耗」统一核算学员剩余课时，支持关键字/班级/授课教师/课时预警筛选与 Excel 导出；该口径同步用于首页统计与学生档案列表。退费新增「退课时数」字段与累计校验，新增退费导出。备份功能升级：备份文件移入 `backups/data/` 专用目录（旧备份自动迁移），新增**一键恢复**（恢复前自动留存快照 + 校验文件有效性 + WAL checkpoint）与**下载到本地**两个能力，备份列表改为切页自动刷新。导入模板新增「已购课时/缴费金额/缴费日期」三列并自动生成初始缴费记录。同步修正文档中多处与实现不符的接口路径（备份、日志、导入）与考勤导出文件名。 |
| v3.3 | 2026-10-01 | **缴费/退费证明图片**：新建/编辑弹窗新增证明图片上传（各最多 5 张，jpg/png/gif/webp，单张 5MB，BLOB 入库），列表新增「缴费证明/退费证明」列显示缩略图、点击即下载原图；删除记录/学生时级联清理图片；顺带修复删除学生不清退费记录的遗留问题。**全站移动端适配**（详见 4.5）：手机端汉堡菜单抽屉导航、弹窗全屏化、表格横滑、选项卡横滑、触控目标加大，电脑端体验零变化。**修复**：用户管理「编辑」按钮调用不存在的函数导致无法编辑的缺陷。 |
| v3.4 | 2026-10-07 | **角色权限体系重构**（详见 2.4）：引入 principal（校长）角色（与 admin 同为唯一角色，均可管理全部功能）；全部权限收敛为权限点集中管理（PERMISSION_MATRIX），前端 can() 与后端 @permission_required 共用一份定义，页面按权限裁剪入口、接口同步拦截；新增权限点 attendance.batch（批量修改考勤）与 course.write（课程写操作，此前接口无权限保护）。**考勤管理**：记录查询页新增「修改查询信息状态」按钮（勾选本次查询结果批量修改状态/备注，权限 attendance.batch）；默认查询范围放宽至一年前并新增「不限日期」快捷查询；只读角色考勤登记页课程信息（教师/上课时间/上课日期/本节课名称）与学员状态纯文本只读。**只读角色限制**：课程管理不可增删改；缴费/退费证明、学生照片、简历附件不可下载（后端拦截 download=1，在线预览不受限）；简历页无导出 PDF/编辑入口，简历编辑页路由加权限拦截。**系统设置**：入口对所有角色开放（页内按权限裁剪，非管理员仅可改密码）；操作日志向教务/教师开放（settings.logs）；修复用户管理面板权限条件包裹缺陷导致的只读页面显示异常。**日志与账户联动**：删除用户时级联删除其全部操作日志；修复登录日志写入时机导致的 user_id 缺失（anonymous 日志）。**登录页**：移除登录按钮下方的默认账号提示文字。 |
