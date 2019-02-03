
/* Drop Tables */

DROP TABLE IF EXISTS child_table;
DROP TABLE IF EXISTS parent_table;




/* Create Tables */

-- 子テーブル
CREATE TABLE child_table
(
	-- ID
	id bigserial NOT NULL,
	-- 親キー
	pid bigint NOT NULL,
	-- 値カラム
	value_column varchar(30),
	PRIMARY KEY (id)
) WITHOUT OIDS;


-- 親テーブル
CREATE TABLE parent_table
(
	-- ID
	id bigserial NOT NULL,
	-- 数値カラム
	number_column bigint,
	-- 固定文字列カラム
	char_column char(5),
	-- 可変長変数カラム
	varchar_column varchar(10),
	-- 日時カラム
	date_column date,
	-- タイムスタンプ_カラム
	timestamp_column timestamp,
	-- NOTNULLカラム
	notnull_column char(1) NOT NULL,
	-- ユニークカラム
	unique_column bigint UNIQUE,
	PRIMARY KEY (id)
) WITHOUT OIDS;



/* Create Foreign Keys */

ALTER TABLE child_table
	ADD FOREIGN KEY (pid)
	REFERENCES parent_table (id)
	ON UPDATE RESTRICT
	ON DELETE RESTRICT
;



/* Comments */

COMMENT ON TABLE child_table IS '子テーブル';
COMMENT ON COLUMN child_table.id IS 'ID';
COMMENT ON COLUMN child_table.pid IS '親キー';
COMMENT ON COLUMN child_table.value_column IS '値カラム';
COMMENT ON TABLE parent_table IS '親テーブル';
COMMENT ON COLUMN parent_table.id IS 'ID';
COMMENT ON COLUMN parent_table.number_column IS '数値カラム';
COMMENT ON COLUMN parent_table.char_column IS '固定文字列カラム';
COMMENT ON COLUMN parent_table.varchar_column IS '可変長変数カラム';
COMMENT ON COLUMN parent_table.date_column IS '日時カラム';
COMMENT ON COLUMN parent_table.timestamp_column IS 'タイムスタンプ_カラム';
COMMENT ON COLUMN parent_table.notnull_column IS 'NOTNULLカラム';
COMMENT ON COLUMN parent_table.unique_column IS 'ユニークカラム';



