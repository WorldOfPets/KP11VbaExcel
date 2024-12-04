--
-- Файл сгенерирован с помощью SQLiteStudio v3.4.4 в Пн сен 16 18:40:44 2024
--
-- Использованная кодировка текста: System
--
PRAGMA foreign_keys = off;
BEGIN TRANSACTION;

-- Таблица: cabinets
CREATE TABLE IF NOT EXISTS cabinets (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name text NOT NULL UNIQUE,
            type text
        );
INSERT INTO cabinets (id, name, type) VALUES (2, '402', NULL);
INSERT INTO cabinets (id, name, type) VALUES (3, '403', NULL);
INSERT INTO cabinets (id, name, type) VALUES (4, '404', NULL);
INSERT INTO cabinets (id, name, type) VALUES (5, '405', NULL);
INSERT INTO cabinets (id, name, type) VALUES (6, '406', NULL);
INSERT INTO cabinets (id, name, type) VALUES (7, '407', NULL);
INSERT INTO cabinets (id, name, type) VALUES (8, '408', NULL);
INSERT INTO cabinets (id, name, type) VALUES (13, '302', NULL);
INSERT INTO cabinets (id, name, type) VALUES (14, '303', NULL);
INSERT INTO cabinets (id, name, type) VALUES (15, '304', NULL);
INSERT INTO cabinets (id, name, type) VALUES (16, '305', NULL);
INSERT INTO cabinets (id, name, type) VALUES (17, '306', NULL);
INSERT INTO cabinets (id, name, type) VALUES (18, '310', NULL);
INSERT INTO cabinets (id, name, type) VALUES (19, '312', NULL);
INSERT INTO cabinets (id, name, type) VALUES (20, '200', NULL);
INSERT INTO cabinets (id, name, type) VALUES (21, '202', NULL);
INSERT INTO cabinets (id, name, type) VALUES (22, '203', NULL);
INSERT INTO cabinets (id, name, type) VALUES (23, '204', NULL);
INSERT INTO cabinets (id, name, type) VALUES (24, '205', NULL);
INSERT INTO cabinets (id, name, type) VALUES (25, '206', NULL);
INSERT INTO cabinets (id, name, type) VALUES (26, '209', NULL);
INSERT INTO cabinets (id, name, type) VALUES (27, '211', NULL);
INSERT INTO cabinets (id, name, type) VALUES (28, '212', NULL);
INSERT INTO cabinets (id, name, type) VALUES (29, '114А', NULL);
INSERT INTO cabinets (id, name, type) VALUES (30, '100', NULL);
INSERT INTO cabinets (id, name, type) VALUES (31, '114', NULL);
INSERT INTO cabinets (id, name, type) VALUES (32, '118', NULL);
INSERT INTO cabinets (id, name, type) VALUES (33, '119', NULL);
INSERT INTO cabinets (id, name, type) VALUES (34, '305А', NULL);

-- Таблица: groups
CREATE TABLE IF NOT EXISTS groups (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name text NOT NULL UNIQUE,
            curator text
        );
INSERT INTO groups (id, name, curator) VALUES (1, 'ИСиП-27 (9кл)', NULL);
INSERT INTO groups (id, name, curator) VALUES (2, 'С-25 (11кл)', NULL);
INSERT INTO groups (id, name, curator) VALUES (3, 'КСиК-22 (9кл)', NULL);
INSERT INTO groups (id, name, curator) VALUES (4, 'ИСиП-33 (9кл)', NULL);
INSERT INTO groups (id, name, curator) VALUES (5, 'ИСиП-32 (9кл)', NULL);
INSERT INTO groups (id, name, curator) VALUES (6, 'С-32 (9кл)', NULL);
INSERT INTO groups (id, name, curator) VALUES (7, 'КСиК-42 (9кл)', NULL);
INSERT INTO groups (id, name, curator) VALUES (8, 'КСиК-11 (9кл)', NULL);
INSERT INTO groups (id, name, curator) VALUES (9, 'ИСиП-13 (11 кл)', NULL);
INSERT INTO groups (id, name, curator) VALUES (10, 'ИСиП-14в (9кл)', NULL);
INSERT INTO groups (id, name, curator) VALUES (11, 'ИСиП-11 (9 кл)', NULL);
INSERT INTO groups (id, name, curator) VALUES (12, 'С-24 (9кл)', NULL);
INSERT INTO groups (id, name, curator) VALUES (13, 'ИСиП-41 (9кл)', NULL);
INSERT INTO groups (id, name, curator) VALUES (14, 'С-12 (9кл)', NULL);
INSERT INTO groups (id, name, curator) VALUES (15, 'КСиК-12 (9кл)', NULL);
INSERT INTO groups (id, name, curator) VALUES (16, 'ИСиП-35 (11кл)', NULL);
INSERT INTO groups (id, name, curator) VALUES (17, 'С-11 (9кл)', NULL);
INSERT INTO groups (id, name, curator) VALUES (18, 'ИСиП-26 (9кл)', NULL);
INSERT INTO groups (id, name, curator) VALUES (19, 'КСиК-31 (9кл)', NULL);
INSERT INTO groups (id, name, curator) VALUES (20, 'С-33 (11кл)', NULL);
INSERT INTO groups (id, name, curator) VALUES (21, 'ИСиП-12 (9 кл)', NULL);
INSERT INTO groups (id, name, curator) VALUES (22, 'С-41 (9кл)', NULL);
INSERT INTO groups (id, name, curator) VALUES (23, 'ИСиП-28 (11кл)', NULL);

-- Таблица: lesson
CREATE TABLE IF NOT EXISTS lesson (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name text NOT NULL UNIQUE
        );
INSERT INTO lesson (id, name) VALUES (2, 'Операционные системы и среды');
INSERT INTO lesson (id, name) VALUES (4, 'Интернет-вещей ');
INSERT INTO lesson (id, name) VALUES (7, 'Моделирование физических процессов в цифровых системах');
INSERT INTO lesson (id, name) VALUES (8, 'Информатика');
INSERT INTO lesson (id, name) VALUES (9, 'Разработка и администрирование систем обработки больших данных (BigData)');
INSERT INTO lesson (id, name) VALUES (10, 'Правовое обеспечение профессиональной деятельности ');
INSERT INTO lesson (id, name) VALUES (11, 'Сертификация информационных систем ');
INSERT INTO lesson (id, name) VALUES (13, 'Информационные технологии ');
INSERT INTO lesson (id, name) VALUES (14, 'Основы дизайн-мышления');
INSERT INTO lesson (id, name) VALUES (15, 'Технология разработки программного обеспечения');
INSERT INTO lesson (id, name) VALUES (16, 'Управление проектами ');
INSERT INTO lesson (id, name) VALUES (17, 'Основы вычислительной техники');
INSERT INTO lesson (id, name) VALUES (18, 'Моделирование и анализ программного обеспечения ');
INSERT INTO lesson (id, name) VALUES (19, 'Компьютерные сети ');
INSERT INTO lesson (id, name) VALUES (20, 'Основы безопасности на производстве ');
INSERT INTO lesson (id, name) VALUES (21, 'Проектный менеджмент в разработке и сопровождении программного обеспечения ');
INSERT INTO lesson (id, name) VALUES (22, 'Основы разработки веб-приложений');
INSERT INTO lesson (id, name) VALUES (23, 'Литература ');
INSERT INTO lesson (id, name) VALUES (25, 'Физические процессы в ЭВМ ');
INSERT INTO lesson (id, name) VALUES (26, 'Основы философии ');
INSERT INTO lesson (id, name) VALUES (27, 'Физическая культура');
INSERT INTO lesson (id, name) VALUES (28, 'Инженерно-техническая поддержка сопровождения информационных систем ');
INSERT INTO lesson (id, name) VALUES (29, 'Инженерная компьютерная графика');
INSERT INTO lesson (id, name) VALUES (30, 'Архитектура аппаратных средств');
INSERT INTO lesson (id, name) VALUES (32, 'Проектирование цифровых устройств ');
INSERT INTO lesson (id, name) VALUES (35, 'Проектирование и разработка интерфейсов пользователя');
INSERT INTO lesson (id, name) VALUES (36, 'История ');
INSERT INTO lesson (id, name) VALUES (37, 'Распределенные системы обработки информации ');
INSERT INTO lesson (id, name) VALUES (39, 'Иностранный язык');
INSERT INTO lesson (id, name) VALUES (40, 'Экономика отрасли ');
INSERT INTO lesson (id, name) VALUES (41, 'Микропроцессорные системы');
INSERT INTO lesson (id, name) VALUES (43, 'Основы  разработки веб - приложений');
INSERT INTO lesson (id, name) VALUES (48, 'Теория вероятностей и математическая статистика ');
INSERT INTO lesson (id, name) VALUES (52, 'Обеспечение безопасности веб-приложений ');
INSERT INTO lesson (id, name) VALUES (57, 'Графический дизайн и мультимедиа ');
INSERT INTO lesson (id, name) VALUES (58, 'Основы проектирования цифровой техники ');
INSERT INTO lesson (id, name) VALUES (59, 'Элементы математической логики ');
INSERT INTO lesson (id, name) VALUES (61, 'Цифровая схемотехника ');
INSERT INTO lesson (id, name) VALUES (67, 'Дискретная математика ');
INSERT INTO lesson (id, name) VALUES (68, 'Прикладная электроника ');
INSERT INTO lesson (id, name) VALUES (69, 'Численные методы ');
INSERT INTO lesson (id, name) VALUES (75, 'Основы электротехники и электронной техники ');
INSERT INTO lesson (id, name) VALUES (76, 'Менеджмент в профессиональной деятельности ');
INSERT INTO lesson (id, name) VALUES (78, 'Основы монтажа');
INSERT INTO lesson (id, name) VALUES (82, 'Биология ');
INSERT INTO lesson (id, name) VALUES (87, 'Беспилотные летательные аппараты ');
INSERT INTO lesson (id, name) VALUES (89, 'Роботизированные системы и комплексы ');
INSERT INTO lesson (id, name) VALUES (90, 'Организация, принципы построения и функционирования компьютерных сетей');
INSERT INTO lesson (id, name) VALUES (95, 'Элементы высшей математики ');
INSERT INTO lesson (id, name) VALUES (96, 'Устройство и функционирование информационной системы ');
INSERT INTO lesson (id, name) VALUES (97, 'Технология выполнения работ по наладке технологического оборудования ');
INSERT INTO lesson (id, name) VALUES (99, 'Обществознание ');
INSERT INTO lesson (id, name) VALUES (101, 'Разговоры о важном');
INSERT INTO lesson (id, name) VALUES (102, 'Безопасность жизнедеятельности ');
INSERT INTO lesson (id, name) VALUES (105, 'Установка и конфигурирование периферийного оборудования ');
INSERT INTO lesson (id, name) VALUES (107, 'Практическая подготовка по технологии выполнения работ по должности служащего "Консультант в области развития цифровой грамотности" ("Цифровой куратор")');
INSERT INTO lesson (id, name) VALUES (109, 'Разработка кода информационных систем ');
INSERT INTO lesson (id, name) VALUES (111, 'Учебная практика ');
INSERT INTO lesson (id, name) VALUES (112, 'Техническое обслуживание и ремонт компьютерных систем и комплексов ');
INSERT INTO lesson (id, name) VALUES (113, 'Химия ');
INSERT INTO lesson (id, name) VALUES (114, 'Разработка приложений для цифровых устройств ');
INSERT INTO lesson (id, name) VALUES (115, 'Проектирование и дизайн информационных систем ');
INSERT INTO lesson (id, name) VALUES (117, 'Основы алгоритмизации и программирования');
INSERT INTO lesson (id, name) VALUES (119, 'Основы проектирования баз данных ');
INSERT INTO lesson (id, name) VALUES (122, 'Технологии физического уровня передачи данных ');
INSERT INTO lesson (id, name) VALUES (127, 'Индивидуальный проект');
INSERT INTO lesson (id, name) VALUES (129, 'Физика ');
INSERT INTO lesson (id, name) VALUES (130, 'Администрирование сетевых операционных систем ');
INSERT INTO lesson (id, name) VALUES (131, 'Проектирование деятельности ');
INSERT INTO lesson (id, name) VALUES (133, 'Программирование микроконтроллеров ');
INSERT INTO lesson (id, name) VALUES (134, 'Организация администрирования компьютерных систем');
INSERT INTO lesson (id, name) VALUES (140, 'Моделирование физических процессов в приложениях');
INSERT INTO lesson (id, name) VALUES (141, 'Психология общения ');
INSERT INTO lesson (id, name) VALUES (142, 'География ');
INSERT INTO lesson (id, name) VALUES (144, 'Математические основы алгоритмизации ');
INSERT INTO lesson (id, name) VALUES (145, 'Стандартизация, сертификация и техническое документоведение ');
INSERT INTO lesson (id, name) VALUES (147, 'Безопасность компьютерных сетей ');
INSERT INTO lesson (id, name) VALUES (149, 'Русский язык ');
INSERT INTO lesson (id, name) VALUES (164, 'Технология работ по профессии рабочего Оператор электронно-вычислительных и вычислительных машин (Проектирование информационных систем на платформе 1С, Проектирование аппаратной части информационных систем)');
INSERT INTO lesson (id, name) VALUES (167, 'Оптимизация веб-приложений  ');
INSERT INTO lesson (id, name) VALUES (189, 'Математика  ');
INSERT INTO lesson (id, name) VALUES (203, 'Инновационные технологии');
INSERT INTO lesson (id, name) VALUES (205, 'Метрология, стандартизация и сертификация ');
INSERT INTO lesson (id, name) VALUES (219, 'Системы виртуализации и контейнеризации  ');
INSERT INTO lesson (id, name) VALUES (246, 'Основы безопасности и защиты Родины ');
INSERT INTO lesson (id, name) VALUES (247, 'Исследовательская и проектная деятельность ');
INSERT INTO lesson (id, name) VALUES (250, 'Разработка мультимедийных приложений ');
INSERT INTO lesson (id, name) VALUES (266, 'Внедрение информационных систем');
INSERT INTO lesson (id, name) VALUES (305, 'Эксплуатация объектов сетевой инфраструктуры ');
INSERT INTO lesson (id, name) VALUES (310, 'Разработка мобильных приложений ');
INSERT INTO lesson (id, name) VALUES (312, 'Иностранный язык в профессиональной деятельности');
INSERT INTO lesson (id, name) VALUES (313, 'Дискретная математика с элементами математической логики');
INSERT INTO lesson (id, name) VALUES (314, 'Администрирование информационных систем (на платформе 1С, Linux, Windows)');

-- Таблица: num
CREATE TABLE IF NOT EXISTS num (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name text NOT NULL UNIQUE,
            timestart text,
            timeend text
        );
INSERT INTO num (id, name, timestart, timeend) VALUES (1, '1', '9:00', '9:45');
INSERT INTO num (id, name, timestart, timeend) VALUES (2, '2', '9:55', '10:40');
INSERT INTO num (id, name, timestart, timeend) VALUES (3, '3', '11:00', '11:45');
INSERT INTO num (id, name, timestart, timeend) VALUES (4, '4', '11:55', '12:40');
INSERT INTO num (id, name, timestart, timeend) VALUES (5, '5', '13:00', '13:45');
INSERT INTO num (id, name, timestart, timeend) VALUES (6, '6', '13:55', '14:40');
INSERT INTO num (id, name, timestart, timeend) VALUES (7, '7', '15:00', '15:45');
INSERT INTO num (id, name, timestart, timeend) VALUES (8, '8', '15:55', '16:40');
INSERT INTO num (id, name, timestart, timeend) VALUES (9, '9', '16:50', '17:35');
INSERT INTO num (id, name, timestart, timeend) VALUES (10, '10', '17:45', '18:30');

-- Таблица: teachers
CREATE TABLE IF NOT EXISTS teachers (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            fio text UNIQUE
        );
INSERT INTO teachers (id, fio) VALUES (1, 'Сапелов В.А.');
INSERT INTO teachers (id, fio) VALUES (2, 'Ливенцева О.А.');
INSERT INTO teachers (id, fio) VALUES (3, 'Поливанов Е.Д.');
INSERT INTO teachers (id, fio) VALUES (4, 'Афиногенов А.А.');
INSERT INTO teachers (id, fio) VALUES (5, 'Коржиновская О.М.');
INSERT INTO teachers (id, fio) VALUES (6, 'Кулицкая Н.С.');
INSERT INTO teachers (id, fio) VALUES (7, 'Воронова Е.А.');
INSERT INTO teachers (id, fio) VALUES (8, 'Колосова Е.Б.');
INSERT INTO teachers (id, fio) VALUES (9, 'Тойвонен Е.Б.');
INSERT INTO teachers (id, fio) VALUES (10, 'Сапелова Е.Н.');
INSERT INTO teachers (id, fio) VALUES (11, 'Пацкевич М.Ю.');
INSERT INTO teachers (id, fio) VALUES (12, 'Моссаковская И.Я.');
INSERT INTO teachers (id, fio) VALUES (13, 'Ночка Е.И.');
INSERT INTO teachers (id, fio) VALUES (14, 'Тимофеева В.М.');
INSERT INTO teachers (id, fio) VALUES (15, 'Чернецов Б.И.');
INSERT INTO teachers (id, fio) VALUES (16, 'Бологов Я.И.');
INSERT INTO teachers (id, fio) VALUES (17, 'Попова К.В.');
INSERT INTO teachers (id, fio) VALUES (18, 'Холопов Р.С.');
INSERT INTO teachers (id, fio) VALUES (19, 'Ластовенко И.Г.');
INSERT INTO teachers (id, fio) VALUES (20, 'Михайлова В.В.');
INSERT INTO teachers (id, fio) VALUES (21, 'Сальникова Н.М.');
INSERT INTO teachers (id, fio) VALUES (23, 'Ильина Е.Ю.');
INSERT INTO teachers (id, fio) VALUES (24, 'Анищенко В.В.');
INSERT INTO teachers (id, fio) VALUES (25, 'Луцков Д.В.');
INSERT INTO teachers (id, fio) VALUES (26, 'Челышева А.С.');
INSERT INTO teachers (id, fio) VALUES (27, 'Вакансия');
INSERT INTO teachers (id, fio) VALUES (28, 'Пояркова Н.Н.');
INSERT INTO teachers (id, fio) VALUES (29, 'Беспалов Д.А.');
INSERT INTO teachers (id, fio) VALUES (30, 'Подстегина А.Д.');
INSERT INTO teachers (id, fio) VALUES (31, 'Каляпин П.О.');
INSERT INTO teachers (id, fio) VALUES (32, 'Барышева С.В.');
INSERT INTO teachers (id, fio) VALUES (33, 'Бочарова Л.А.');
INSERT INTO teachers (id, fio) VALUES (34, 'Виноградов В.В.');
INSERT INTO teachers (id, fio) VALUES (35, 'Строкань О.О.');
INSERT INTO teachers (id, fio) VALUES (36, 'Сыроегин В.В.');
INSERT INTO teachers (id, fio) VALUES (37, 'Шаврова О.С.');
INSERT INTO teachers (id, fio) VALUES (38, 'Королева Н.В.');
INSERT INTO teachers (id, fio) VALUES (39, 'Варов Д.Е.');
INSERT INTO teachers (id, fio) VALUES (40, 'Авдеева Е.А.');
INSERT INTO teachers (id, fio) VALUES (41, 'Маслов В.В.');
INSERT INTO teachers (id, fio) VALUES (42, 'Володина И.В.');
INSERT INTO teachers (id, fio) VALUES (43, 'Серегина Л.Д.');
INSERT INTO teachers (id, fio) VALUES (44, 'Хохлов К.В.');
INSERT INTO teachers (id, fio) VALUES (45, 'Виноградова С.В.');
INSERT INTO teachers (id, fio) VALUES (46, 'Куратор группы');

COMMIT TRANSACTION;
PRAGMA foreign_keys = on;
