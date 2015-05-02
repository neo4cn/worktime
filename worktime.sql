--DDL for worktime at sqlite3

CREATE TABLE `worktime` (
	`date`	TEXT,
	`personid`	TEXT,
	`projectid`	TEXT,
	`phaseid`	TEXT NOT NULL,
	`hour`	INTEGER NOT NULL,
	`overflag`	INTEGER NOT NULL,
	`remark`	TEXT,
	PRIMARY KEY(date,personid,projectid,phaseid,overflag)
);

CREATE TABLE `person` (
	`id`	TEXT,
	`name`	TEXT NOT NULL,
	PRIMARY KEY(id)
);

CREATE TABLE `project` (
	`id`	TEXT,
	`name`	TEXT NOT NULL,
	PRIMARY KEY(id)
);

CREATE TABLE `phase` (
	`id`	TEXT,
	`name`	TEXT NOT NULL,
	PRIMARY KEY(id)
);
