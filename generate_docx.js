const {
  Document, Packer, Paragraph, TextRun, AlignmentType,
  HeadingLevel, Table, TableRow, TableCell, WidthType,
  BorderStyle, PageBreak, TabStopPosition, TabStopType,
  SectionType, Header, Footer, NumberFormat, convertInchesToTwip
} = require('docx');
const fs = require('fs');

// ====== 排版常量 ======
const FONT_BODY = 'SimSun';        // 宋体
const FONT_HEADING = 'SimHei';     // 黑体
const SIZE_XIAOSI = 24;            // 小四 = 12pt = 24半点
const SIZE_HAO3 = 32;              // 三号 = 16pt = 32半点
const SIZE_HAO4 = 28;              // 四号 = 14pt = 28半点
const SIZE_HAO2 = 44;              // 二号 = 22pt (封面标题)
const SIZE_HAO3_COVER = 32;        // 三号 (封面信息)
const LINE_SPACING_15 = 360;       // 1.5倍行距 = 360/240
const INDENT_2CHAR = 480;          // 首行缩进2字符 = 12pt*2*20 = 480 twips

// ====== 辅助函数 ======

/** 正文段落: 宋体小四, 1.5倍行距, 首行缩进2字符 */
function bodyPara(text) {
  return new Paragraph({
    spacing: { line: LINE_SPACING_15, lineRule: 'auto' },
    indent: { firstLine: INDENT_2CHAR },
    children: [new TextRun({ text, font: FONT_BODY, size: SIZE_XIAOSI })]
  });
}

/** 正文段落 + 加粗文本 */
function bodyParaBold(boldText, normalText) {
  return new Paragraph({
    spacing: { line: LINE_SPACING_15, lineRule: 'auto' },
    indent: { firstLine: INDENT_2CHAR },
    children: [
      new TextRun({ text: boldText, font: FONT_BODY, size: SIZE_XIAOSI, bold: true }),
      new TextRun({ text: normalText, font: FONT_BODY, size: SIZE_XIAOSI })
    ]
  });
}

/** 一级标题: 三号黑体居中 */
function h1(text) {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 300, after: 200, line: LINE_SPACING_15, lineRule: 'auto' },
    children: [new TextRun({ text, font: FONT_HEADING, size: SIZE_HAO3, bold: true })]
  });
}

/** 二级标题: 四号黑体左对齐 */
function h2(text) {
  return new Paragraph({
    alignment: AlignmentType.LEFT,
    spacing: { before: 200, after: 100, line: LINE_SPACING_15, lineRule: 'auto' },
    children: [new TextRun({ text, font: FONT_HEADING, size: SIZE_HAO4, bold: true })]
  });
}

/** 空行 */
function emptyLine() {
  return new Paragraph({
    spacing: { line: LINE_SPACING_15, lineRule: 'auto' },
    children: [new TextRun({ text: '', font: FONT_BODY, size: SIZE_XIAOSI })]
  });
}

/** 居中文本 */
function centerText(text, font, size) {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { line: LINE_SPACING_15, lineRule: 'auto' },
    children: [new TextRun({ text, font: font || FONT_BODY, size: size || SIZE_XIAOSI })]
  });
}

/** 表格单元格 */
function cell(text, width, bold) {
  return new TableCell({
    width: { size: width, type: WidthType.DXA },
    children: [new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new TextRun({ text, font: FONT_BODY, size: SIZE_XIAOSI, bold })],
    })]
  });
}

// ====== 构建文档 ======
const doc = new Document({
  styles: {
    default: {
      document: {
        run: { font: FONT_BODY, size: SIZE_XIAOSI },
        paragraph: { spacing: { line: LINE_SPACING_15, lineRule: 'auto' } }
      }
    }
  },
  sections: [
    // ========== 封面 ==========
    {
      properties: {
        page: { margin: { top: 1200, bottom: 1200, left: 1440, right: 1440 } }
      },
      children: [
        emptyLine(), emptyLine(), emptyLine(),

        // 学校/课程标题
        centerText('面向对象程序设计 (Java) 课程设计报告', FONT_HEADING, SIZE_HAO2),

        emptyLine(), emptyLine(), emptyLine(),

        // 题目
        centerText('题目：基于网络的彩票购买抽奖系统', FONT_HEADING, SIZE_HAO3_COVER),

        emptyLine(), emptyLine(),

        // 信息表格
        new Table({
          width: { size: 100, type: WidthType.PERCENTAGE },
          rows: [
            rowData('成绩等级', 'A'),
            rowData('姓　　名', '陶震'),
            rowData('学　　号', '202503013135'),
            rowData('完成人数', '1 人'),
            rowData('开发技术', 'Java Socket + 多线程 + 对象序列化持久化'),
            rowData('完成日期', '2026 年 06 月 29 日'),
          ]
        }),

        emptyLine(), emptyLine(), emptyLine(),

        // 摘要
        h2('摘要'),
        bodyPara(
          '本次课程设计采用 Java 语言开发 C/S 架构网络彩票抽奖系统，基于 Socket 完成客户端与服务器网络通信，' +
          '使用线程池实现多客户端并发处理，通过 Java 对象序列化技术将用户数据、购彩记录、开奖信息持久化存储至本地文件。' +
          '系统分为服务端管理程序与控制台客户端，完整实现双色球全套业务流程。'
        ),
        bodyPara(
          '普通用户可完成账号注册、登录、账户充值、自选号码购彩、查询历史购彩记录、接收中奖推送通知；' +
          '管理员在服务器控制台可执行手动摇奖、当期销售数据统计操作。' +
          '项目采用 ConcurrentHashMap、CopyOnWriteArrayList 结合同步锁保障多线程并发下的数据安全，' +
          '严格遵循双色球选号规则与六档中奖奖金标准。' +
          '整体代码遵循面向对象设计思想，模块划分独立清晰，满足单一职责、封装、开闭等设计原则。' +
          '测试阶段可同时启动 3 个及以上客户端并发操作，服务器重启后数据不丢失，全部功能与技术要求均实现完成。'
        ),

        emptyLine(),
        bodyParaBold('关键词：', 'Java；Socket；多线程；C/S 架构；对象序列化；双色球系统'),

        // ====== 目录 ======
        new Paragraph({ children: [new PageBreak()] }),
        h1('目录'),
        emptyLine(),
        tocItem('1 题目概述', '1'),
        tocItem('1.1 项目基础信息', '1'),
        tocItem('1.2 整体技术说明', '1'),
        tocItem('2 功能需求', '2'),
        tocItem('2.1 服务器端功能需求', '2'),
        tocItem('2.2 客户端功能需求', '2'),
        tocItem('2.3 技术硬性要求', '3'),
        tocItem('3 需求分析', '4'),
        tocItem('3.1 系统角色划分', '4'),
        tocItem('3.2 双色球彩票规则', '4'),
        tocItem('3.3 中奖奖金判定规则', '4'),
        tocItem('3.4 Socket 通信协议设计', '5'),
        tocItem('3.5 并发处理实现方案', '5'),
        tocItem('4 系统概要设计', '6'),
        tocItem('4.1 系统整体 C/S 架构', '6'),
        tocItem('4.2 项目包与类层次结构', '7'),
        tocItem('4.3 核心设计决策', '8'),
        tocItem('5 系统详细设计', '9'),
        tocItem('5.1 Message 通信消息类', '9'),
        tocItem('5.2 UserManager 用户管理模块', '10'),
        tocItem('5.3 LotteryManager 彩票业务模块', '11'),
        tocItem('5.4 DrawEngine 摇奖与中奖判定引擎', '12'),
        tocItem('5.5 ClientHandler 客户端处理线程', '13'),
        tocItem('5.6 LotteryClient 客户端主程序', '13'),
        tocItem('6 系统测试方案', '14'),
        tocItem('6.1 测试运行环境', '14'),
        tocItem('6.2 功能测试用例', '14'),
        tocItem('6.3 编译运行步骤', '16'),
        tocItem('7 设计总结', '17'),
        tocItem('7.1 系统核心技术亮点', '17'),
        tocItem('7.2 面向对象设计原则体现', '18'),
        tocItem('7.3 系统现存不足与优化方向', '18'),
        tocItem('8 项目文件清单', '20'),
        tocItem('参考文献', '21'),

        // ====== 正文 ======
        new Paragraph({ children: [new PageBreak()] }),

        // ---- 第1章 ----
        h1('1 题目概述'),
        h2('1.1 项目基础信息'),
        bodyParaBold('项目名称：', '基于网络的彩票购买抽奖程序'),
        bodyParaBold('开发人员：', '陶震'),
        bodyParaBold('学号：', '202503013135'),
        bodyParaBold('项目评级：', 'A 级单人课程设计'),
        bodyParaBold('技术栈：', 'Java Socket 网络编程、固定线程池、Java 对象序列化持久化'),
        bodyParaBold('架构模式：', '客户端 / 服务器（C/S）分布式架构'),
        bodyParaBold('交互界面：', '控制台字符交互，模拟线上彩票 APP 操作页面'),

        h2('1.2 整体技术说明'),
        bodyPara(
          '服务器监听本地 8888 端口，主线程循环接收客户端 Socket 连接请求，通过固定线程池为每一个客户端分配独立处理线程，' +
          '支持多用户同时在线操作，互不阻塞。客户端通过 Socket 与服务端建立长连接，所有请求、响应统一封装为可序列化 Message 对象传输，简化通信逻辑。'
        ),
        bodyPara(
          '用户账户信息、购彩订单、每期开奖结果通过对象序列化写入 data 目录下.dat 文件持久化存储，' +
          '服务器关闭重启后可自动读取全部历史数据，保证数据不丢失。'
        ),

        // ---- 第2章 ----
        h1('2 功能需求'),
        h2('2.1 服务器端功能需求'),
        bodyPara('用户注册：接收客户端注册请求，校验用户名唯一性，持久化保存用户名、密码、账户余额。'),
        bodyPara('购彩记录存储：永久保存每一笔购彩订单，记录彩票期号、自选号码、购买时间、中奖状态。'),
        bodyPara('中奖通知推送：用户登录时自动查询未读中奖记录，实时推送中奖提醒。'),
        bodyPara('管理员摇奖功能：服务器控制台输入指令触发摇奖，随机生成当期双色球开奖号码。'),
        bodyPara('销售数据统计：按期统计当期彩票总购买注数、总销售金额并打印展示。'),

        h2('2.2 客户端功能需求（模拟手机彩票客户端）'),
        bodyPara('用户登录：输入账号密码登录系统，登录成功自动展示账户余额与未读中奖通知。'),
        bodyPara('账户充值：自定义充值金额，服务端更新余额并返回最新账户金额。'),
        bodyPara('双色球购彩：用户手动选择 6 个不重复 1~33 红球、1 个 1~16 蓝球，每注自动扣除 2 元。'),
        bodyPara('历史购彩记录：查看个人全部购彩订单，展示每笔订单号码、购买时间、中奖金额。'),
        bodyPara('中奖提醒：登录后自动弹出所有未读中奖通知，通知读取后不再重复推送。'),

        h2('2.3 技术硬性要求'),
        bodyPara('服务器采用多线程机制，支持多个客户端并发操作，线程间互不阻塞。'),
        bodyPara('程序支持至少 3 个客户端同时启动、在线并发测试。'),
        bodyPara('所有实体类、消息传输类实现 Serializable 序列化接口，支持网络传输与本地文件持久化。'),

        // ---- 第3章 ----
        h1('3 需求分析'),
        h2('3.1 系统角色划分'),
        bodyPara('普通用户（客户端使用者）：注册账号、登录、充值、购买彩票、查询历史订单、接收中奖消息。'),
        bodyPara('系统管理员（服务器控制台操作者）：执行摇奖、查看当期销售统计、关闭服务器程序。'),

        h2('3.2 双色球彩票规则'),
        bodyPara('红球取值范围：数字 1~33，单注选择 6 个不重复数字，系统自动升序排序。'),
        bodyPara('蓝球取值范围：数字 1~16，单注仅选择 1 个数字。'),
        bodyPara('单注售价：人民币 2 元。'),

        h2('3.3 中奖奖金判定规则'),
        emptyLine(),
        new Table({
          width: { size: 100, type: WidthType.PERCENTAGE },
          rows: [
            tableHeaderRow(['中奖等级', '匹配条件', '单注奖金']),
            tableRow(['一等奖', '6 红 + 1 蓝全匹配', '5,000,000 元']),
            tableRow(['二等奖', '6 红匹配，蓝球不匹配', '200,000 元']),
            tableRow(['三等奖', '5 红 + 1 蓝匹配', '3,000 元']),
            tableRow(['四等奖', '5 红无蓝 / 4 红 + 1 蓝', '200 元']),
            tableRow(['五等奖', '4 红无蓝 / 3 红 + 1 蓝', '10 元']),
            tableRow(['六等奖', '仅蓝球匹配', '5 元']),
          ]
        }),
        emptyLine(),

        h2('3.4 Socket 通信协议设计'),
        bodyPara(
          '自定义 Message 消息实体封装全部网络交互数据，通过枚举 MessageType 区分消息类型：'
        ),
        bodyPara('请求类型：REGISTER 注册、LOGIN 登录、RECHARGE 充值、BUY_LOTTERY 购彩、MY_RECORDS 查看记录、DRAW 摇奖、STATISTICS 销售统计、EXIT 退出。'),
        bodyPara('响应类型：SUCCESS 操作成功、ERROR 操作失败、WIN_NOTIFY 中奖通知。'),

        h2('3.5 并发处理实现方案'),
        bodyPara(
          '服务端采用 ThreadPoolExecutor 固定线程池，最大并发线程 50，每个客户端连接分配独立 ClientHandler 线程；' +
          '数据容器采用线程安全集合：ConcurrentHashMap 存储用户信息、CopyOnWriteArrayList 存储购彩记录；' +
          '核心业务方法添加 synchronized 同步锁，防止多线程同时读写造成数据错乱。'
        ),

        // ---- 第4章 ----
        h1('4 系统概要设计'),
        h2('4.1 系统整体 C/S 架构'),
        bodyPara('系统采用经典 C/S 三层架构：客户端展示层 → 服务端业务逻辑层 → 文件持久化层。'),
        bodyPara('服务端 LotteryServer 主线程循环调用 ServerSocket.accept() 监听 8888 端口，接收到连接后提交至线程池；每个客户端由独立 ClientHandler 线程服务，调用 UserManager、LotteryManager 执行业务逻辑，最终通过 Java 序列化写入 data 目录 .dat 文件持久化。'),
        emptyLine(),
        bodyParaBold('图 4.1', ' 系统整体 C/S 分层架构图（详见程序源代码结构）'),
        emptyLine(),

        h2('4.2 项目包与类层次结构'),
        bodyParaBold('（1）common 公共包（通用序列化实体、常量）', ''),
        bodyPara('Message：网络通信消息封装类，实现序列化。'),
        bodyPara('LotteryConstants：全局系统常量、文件路径、彩票规则参数。'),
        bodyParaBold('（2）server 服务端包', ''),
        bodyPara('LotteryServer：服务器启动主类，端口监听、线程池初始化、管理员控制台'),
        bodyPara('ClientHandler：单客户端处理线程，解析消息分发业务逻辑'),
        bodyPara('UserManager：用户账户管理，注册、登录、充值、扣款、持久化读写'),
        bodyPara('User：用户实体类（用户名、密码、余额、未读中奖通知）'),
        bodyPara('LotteryManager：彩票业务总管理，购彩、开奖、查询订单、销售统计'),
        bodyPara('LotteryRecord：购彩订单实体（期号、号码、购买时间、中奖状态）'),
        bodyPara('DrawResult：每期开奖结果实体（期号、开奖红球、蓝球、开奖时间）'),
        bodyPara('DrawEngine：摇奖算法引擎，随机生成号码、中奖等级匹配计算'),
        bodyParaBold('（3）client 客户端包', ''),
        bodyPara('LotteryClient：客户端主程序，Socket 连接、控制台菜单交互、消息发送接收'),

        h2('4.3 核心设计决策'),
        bodyPara('线程池方案：固定大小线程池，限制最大并发 50，避免无限创建线程耗尽服务器资源。'),
        bodyPara('线程安全：使用 JUC 并发容器 + 同步方法，多用户同时读写账户、订单无脏数据。'),
        bodyPara('持久化方案：Java 原生 Object 序列化，程序启动自动加载 data 目录数据，修改后自动保存。'),
        bodyPara('通信方案：统一 Message 对象传输，无需拼接字符串，扩展新功能仅增加枚举类型即可。'),
        bodyPara('中奖通知机制：登录拉取模式，用户登录时查询全部未读通知，读取完成后清除，避免重复提醒。'),

        // ---- 第5章 ----
        h1('5 系统详细设计'),
        h2('5.1 Message 通信消息类'),
        bodyParaBold('核心字段', '：'),
        bodyPara('type：MessageType 枚举，标记当前消息请求 / 响应类型'),
        bodyPara('username、password：用户身份认证信息'),
        bodyPara('amount：double 金额，用于充值扣款'),
        bodyPara('numbers：int 数组，存储用户选择 6 个红球'),
        bodyPara('blueNumber：int，用户选择蓝球号码'),
        bodyPara('content：通用文本提示信息'),
        bodyPara('data：Map<String, Object> 扩展字段，传递余额、订单列表、通知等复杂数据'),
        emptyLine(),
        bodyParaBold('标准消息交互流程', '：客户端向服务器发送请求，服务器处理后返回 SUCCESS/ERROR 响应。' +
          '注册请求携带 username 和 password，登录成功后返回余额与中奖通知，' +
          '购彩请求携带红球数组和蓝球号码，服务器扣款后返回订单记录与最新余额。'),
        emptyLine(),

        h2('5.2 UserManager 用户管理模块'),
        bodyParaBold('核心数据结构：', 'ConcurrentHashMap<String, User>，key 为用户名，value 为用户实体。'),
        bodyPara('register(String username, String password)：校验用户名重复，创建用户对象并持久化。'),
        bodyPara('login(String username, String password)：校验账号密码，返回登录结果。'),
        bodyPara('recharge(String username, double amount)：增加账户余额，自动保存文件。'),
        bodyPara('deduct(String username, double amount)：扣除 2 元购彩费用，余额不足返回 false。'),
        bodyPara('load() / save()：序列化读取 / 写入 users.dat 持久化文件。'),

        h2('5.3 LotteryManager 彩票业务模块'),
        bodyPara('buy(String username, int[] reds, int blue)：校验号码合法性，扣款生成订单记录。'),
        bodyPara('performDraw()：执行摇奖，生成当期开奖号，自动匹配所有订单生成中奖通知。'),
        bodyPara('fetchNotifications(String username)：获取用户未读中奖消息，读取后清空。'),
        bodyPara('getUserRecords(String username)：查询该用户全部历史购彩订单。'),
        bodyPara('getStatistics()：统计当期总购彩注数、总销售金额，格式化输出统计文本。'),
        bodyParaBold('期号生成规则：', 'yyyyMMdd-NNN，例如 20260629-001。'),

        h2('5.4 DrawEngine 摇奖与中奖判定引擎'),
        bodyPara(
          'draw() 摇奖方法：使用 TreeSet 自动排序，随机生成 6 个 1~33 不重复红球，搭配 1 个 1~16 蓝球。'
        ),
        bodyPara(
          'checkPrize() 判定方法：利用 HashSet 求用户红球与开奖红球交集，统计匹配红球数量，' +
          '结合蓝球是否匹配，对照规则表判定中奖等级与奖金。'
        ),

        h2('5.5 ClientHandler 客户端处理线程'),
        bodyPara('创建 ObjectInputStream、ObjectOutputStream 处理序列化流；'),
        bodyPara('循环读取客户端发送的 Message 请求；'),
        bodyPara('根据 MessageType 枚举类型分发至对应业务处理方法；'),
        bodyPara('调用 UserManager、LotteryManager 执行业务逻辑；'),
        bodyPara('封装响应 Message 对象发送回客户端；'),
        bodyPara('捕获 EOFException 异常，处理客户端主动断开连接。'),

        h2('5.6 LotteryClient 客户端主程序'),
        bodyPara('未登录状态 → 展示注册 / 登录选择菜单；'),
        bodyPara('已登录状态 → 主菜单（充值、购彩、查看记录、查看中奖通知、退出）；'),
        bodyPara('connect(host, port)：建立 Socket 网络连接；'),
        bodyPara('sendMessage(request)：发送消息并同步阻塞等待服务器响应；'),
        bodyPara('parseNumbers()：解析并校验用户输入的红球、蓝球号码合法性；'),
        bodyPara('doBuyLottery()：选号确认、发送购彩请求、展示购买结果。'),

        // ---- 第6章 ----
        h1('6 系统测试方案'),
        h2('6.1 测试运行环境'),
        bodyPara('运行环境：JDK 1.8 及以上版本'),
        bodyPara('启动方式：同一电脑多终端窗口，1 个服务器、3 个客户端同时运行'),
        bodyPara('编译命令：javac -encoding UTF-8 -d bin src/common/*.java src/server/*.java src/client/*.java'),
        bodyPara('启动服务器：java -cp bin server.LotteryServer'),
        bodyPara('启动客户端：java -cp bin client.LotteryClient'),

        h2('6.2 功能测试用例'),
        emptyLine(),
        bodyParaBold('TC-01 用户注册', '：客户端选择注册，输入用户名、密码。预期：全新用户名注册成功；重复用户名提示注册失败。'),
        bodyParaBold('TC-02 账户余额充值', '：登录账号，输入充值 100 元。预期：账户余额更新为 100 元，返回充值成功提示。'),
        bodyParaBold('TC-03 余额充足购买彩票', '：账户余额 ≥ 2 元，自选合法号码购彩。预期：账户扣除 2 元，生成购彩记录，展示订单信息。'),
        bodyParaBold('TC-04 余额不足购买彩票', '：账户余额为 0，发起购彩请求。预期：服务器返回错误提示"余额不足，请先充值"。'),
        bodyParaBold('TC-05 多客户端并发购彩', '：同时开启 3 个客户端，同步执行购彩操作。预期：全部客户端操作正常完成，用户余额、订单数据无错乱。'),
        bodyParaBold('TC-06 管理员摇奖与中奖通知', '：服务器控制台执行 draw 摇奖指令；中奖用户重新登录客户端。预期：登录后自动弹出中奖通知，已读通知不再重复展示。'),
        bodyParaBold('TC-07 个人购彩记录查询', '：多次购彩后选择查看记录功能。预期：完整展示所有历史订单，标注每笔订单中奖状态与奖金。'),
        bodyParaBold('TC-08 当期销售统计', '：服务器控制台输入 stats 统计指令。预期：控制台打印当期总购彩注数、总销售金额。'),
        bodyParaBold('TC-09 数据持久化验证', '：完成购彩、充值操作 → 关闭服务器 → 重新启动服务器、登录账号。预期：用户余额、购彩记录、开奖数据全部保留，无丢失。'),

        h2('6.3 编译运行步骤'),
        bodyPara('① 在项目根目录打开命令行，执行编译指令，生成 bin 字节码文件夹；'),
        bodyPara('② 新开终端窗口，运行服务器主程序，服务器启动后监听 8888 端口；'),
        bodyPara('③ 开启多个新终端，分别运行客户端程序，输入账号登录操作；'),
        bodyPara('④ 服务器控制台支持三条管理员指令：draw（摇奖）、stats（统计）、exit（关闭服务）。'),

        // ---- 第7章 ----
        h1('7 设计总结'),
        h2('7.1 系统核心技术亮点'),
        bodyPara('C/S 分布式 Socket 架构：客户端与服务器网络分离，支持多终端远程访问程序；'),
        bodyPara('线程池并发处理：复用线程资源，支持大量客户端同时在线，资源开销更低；'),
        bodyPara('完整线程安全机制：并发容器 + 同步锁，解决多线程读写数据冲突问题；'),
        bodyPara('对象序列化持久化：无需数据库，本地文件存储全部业务数据，重启不丢失；'),
        bodyPara('统一消息通信封装：Message 统一封装所有交互数据，拓展功能简单便捷；'),
        bodyPara('登录拉取式中奖通知：自动推送未读消息，用户体验清晰简洁。'),

        h2('7.2 面向对象设计原则体现'),
        bodyPara('单一职责原则：每个类职责独立，UserManager 只管理用户、DrawEngine 只处理摇奖逻辑；'),
        bodyPara('开闭原则：新增消息类型、彩票玩法仅需扩展枚举，无需修改原有核心业务代码；'),
        bodyPara('依赖倒置：业务处理类依赖通用实体，不绑定具体实现；'),
        bodyPara('封装特性：所有实体类字段私有，仅通过 getter/setter 访问，屏蔽内部数据细节。'),

        h2('7.3 系统现存不足与优化改进方向'),
        bodyPara('界面限制：当前仅控制台文字交互，后续可扩展 Swing/JavaFX 图形可视化界面；'),
        bodyPara('账号安全：密码明文序列化存储，可引入 BCrypt 哈希加密存储密码；'),
        bodyPara('存储方式：文件序列化查询效率低，可替换 SQLite/MySQL 数据库存储数据；'),
        bodyPara('日志管理：仅使用 System.out 打印信息，可接入 Log4j 实现分级日志记录；'),
        bodyPara('IO 性能：当前 BIO 阻塞 IO，高并发场景可升级 NIO 非阻塞网络模型提升性能。'),

        // ---- 第8章 ----
        h1('8 项目文件清单'),
        bodyParaBold('项目根目录：', 'LotterySystem/'),
        bodyPara('├── src/'),
        bodyPara('│   ├── common/'),
        bodyPara('│   │   ├── Message.java —— 通信消息协议类'),
        bodyPara('│   │   └── LotteryConstants.java —— 全局常量配置'),
        bodyPara('│   ├── server/'),
        bodyPara('│   │   ├── LotteryServer.java —— 服务器主程序'),
        bodyPara('│   │   ├── ClientHandler.java —— 客户端处理线程'),
        bodyPara('│   │   ├── UserManager.java —— 用户管理模块'),
        bodyPara('│   │   ├── User.java —— 用户实体类'),
        bodyPara('│   │   ├── LotteryManager.java —— 彩票业务管理'),
        bodyPara('│   │   ├── LotteryRecord.java —— 购彩记录实体'),
        bodyPara('│   │   ├── DrawResult.java —— 开奖结果实体'),
        bodyPara('│   │   └── DrawEngine.java —— 摇奖判定引擎'),
        bodyPara('│   └── client/'),
        bodyPara('│       └── LotteryClient.java —— 客户端主程序'),
        bodyPara('├── 课程设计报告.docx —— 本设计文档'),
        bodyPara('└── README.txt —— 项目运行说明'),
        emptyLine(),
        bodyParaBold('项目源码总计：', '11 个 Java 源文件，代码总量约 1200 行。'),

        // ---- 参考文献 ----
        new Paragraph({ children: [new PageBreak()] }),
        h1('参考文献'),
        emptyLine(),
        bodyPara('[1] 耿祥义，张跃平. Java 面向对象程序设计 [M]. 清华大学出版社，2020.'),
        bodyPara('[2] 陈雄华. Java 网络编程实战 [M]. 机械工业出版社，2021.'),
        bodyPara('[3] 李刚. 疯狂 Java 讲义 [M]. 电子工业出版社，2019.'),
        bodyPara('[4] Oracle 官方文档. Java IO 与对象序列化 API 开发手册.'),
        bodyPara('[5] Oracle 官方文档. Java 多线程并发编程线程池使用指南.'),
      ]
    }
  ]
});

// ====== 生成文件 ======
Packer.toBuffer(doc).then(buffer => {
  const outPath = 'C:\\Users\\user\\Desktop\\《Java 课程设计 - 彩票系统 - 陶震》.docx';
  fs.writeFileSync(outPath, buffer);
  console.log('✅ 文档生成成功: ' + outPath);
  console.log('文件大小: ' + (buffer.length / 1024).toFixed(1) + ' KB');
}).catch(err => {
  console.error('生成失败: ', err);
});

// ====== 辅助表格函数 ======
function tocItem(text, page) {
  return new Paragraph({
    spacing: { line: 400, lineRule: 'auto' },
    children: [new TextRun({ text, font: FONT_BODY, size: SIZE_XIAOSI })],
    tabStops: [{
      type: TabStopType.RIGHT,
      position: 8500
    }, {
      type: TabStopType.LEFT,
      position: 8501
    }]
  });
}

function rowData(label, value) {
  return new TableRow({
    children: [
      new TableCell({
        width: { size: 2500, type: WidthType.DXA },
        children: [new Paragraph({
          alignment: AlignmentType.RIGHT,
          children: [new TextRun({ text: label + '：', font: FONT_HEADING, size: SIZE_XIAOSI, bold: true })]
        })]
      }),
      new TableCell({
        width: { size: 6000, type: WidthType.DXA },
        children: [new Paragraph({
          alignment: AlignmentType.LEFT,
          children: [new TextRun({ text: value, font: FONT_BODY, size: SIZE_XIAOSI })]
        })]
      })
    ]
  });
}

function tableHeaderRow(headers) {
  return new TableRow({
    tableHeader: true,
    children: headers.map(h => new TableCell({
      width: { size: 3000, type: WidthType.DXA },
      shading: { fill: 'D9E2F3' },
      children: [new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: h, font: FONT_HEADING, size: SIZE_XIAOSI, bold: true })]
      })]
    }))
  });
}

function tableRow(cells) {
  return new TableRow({
    children: cells.map(c => new TableCell({
      width: { size: 3000, type: WidthType.DXA },
      children: [new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: c, font: FONT_BODY, size: SIZE_XIAOSI })]
      })]
    }))
  });
}
