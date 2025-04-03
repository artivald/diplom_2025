BEGIN TRANSACTION;
CREATE TABLE IF NOT EXISTS "Resourcses" (
	"id"	INTEGER,
	"name"	TEXT NOT NULL,
	"postfix"	TEXT,
	PRIMARY KEY("id")
);
CREATE TABLE IF NOT EXISTS "Access" (
	"id"	INTEGER,
	"id_user"	INTEGER NOT NULL,
	"id_res"	INTEGER NOT NULL,
	PRIMARY KEY("id"),
	FOREIGN KEY("id_user") REFERENCES "user"("id")
);
CREATE TABLE IF NOT EXISTS "user" (
	"id"	INTEGER,
	"surname"	TEXT NOT NULL,
	"name"	TEXT NOT NULL,
	"patronymic"	TEXT,
	"login"	TEXT NOT NULL,
	"password"	TEXT NOT NULL,
	"division"	TEXT,
	"post"	TEXT,
	"faculty"	TEXT,
	PRIMARY KEY("id")
);
CREATE TABLE IF NOT EXISTS "division" (
	"Id_division"	INTEGER,
	"division"	TEXT,
	PRIMARY KEY("Id_division")
);
CREATE TABLE IF NOT EXISTS "post" (
	"id_post"	INTEGER,
	"post"	TEXT,
	PRIMARY KEY("id_post")
);
CREATE TABLE IF NOT EXISTS "Faculties" (
	"facultie_short"	TEXT,
	"facultie_full"	TEXT
);
COMMIT;
