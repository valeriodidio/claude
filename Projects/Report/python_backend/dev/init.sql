-- Schema della tabella actual_turnover per test in locale.
-- Viene eseguito automaticamente da docker-compose al primo avvio del container MySQL.

CREATE DATABASE IF NOT EXISTS yeppon_stats
    DEFAULT CHARACTER SET utf8mb4
    COLLATE utf8mb4_unicode_ci;

USE yeppon_stats;

CREATE TABLE IF NOT EXISTS actual_turnover (
    id               BIGINT AUTO_INCREMENT PRIMARY KEY,
    marca            VARCHAR(50),
    quantita         INT,
    turnover         DOUBLE,
    prezzo_acquisto  DOUBLE,
    macrocat         VARCHAR(50),
    cate1            VARCHAR(100),
    cate2            VARCHAR(100),
    cate3            VARCHAR(100),
    nomemktp         VARCHAR(50),
    dataordine       TIMESTAMP,
    provincia        VARCHAR(20),
    codice           VARCHAR(30),
    id_p             INT,
    nome             VARCHAR(300),
    mktpordinato     VARCHAR(50),
    order_id         INT,
    fornitore        VARCHAR(50),
    ean              VARCHAR(50),
    user_type        VARCHAR(30),
    cod_fornitore    VARCHAR(50),
    tipo_spedizione  VARCHAR(150),
    importo_spe      DOUBLE,
    INDEX idx_dataordine (dataordine),
    INDEX idx_macrocat   (macrocat),
    INDEX idx_nomemktp   (nomemktp),
    INDEX idx_codice     (codice)
);

-- Database smart2: contiene ordini_cliente (usato in join per riconoscere H&A)
CREATE DATABASE IF NOT EXISTS smart2
    DEFAULT CHARACTER SET utf8mb4
    COLLATE utf8mb4_unicode_ci;

USE smart2;

CREATE TABLE IF NOT EXISTS ordini_cliente (
    id        INT PRIMARY KEY,
    idUtente  INT NOT NULL,
    INDEX idx_id_utente (idUtente)
);

USE yeppon_stats;

-- Utente di sola lettura usato dall'app
CREATE USER IF NOT EXISTS 'usr_report'@'%' IDENTIFIED BY 'dev_password_change_me';
GRANT SELECT ON yeppon_stats.actual_turnover  TO 'usr_report'@'%';
GRANT SELECT ON smart2.ordini_cliente         TO 'usr_report'@'%';
FLUSH PRIVILEGES;
