CREATE TABLE Hymn(
    HymnName    VARCHAR(255)             NOT NULL,
    Version     UNSIGNED TINYINT         NOT NULL,
    Number      UNSIGNED TINYINT         NOT NULL,
    End         UNSIGNED TINYINT         NOT NULL,
    Lyrics      TEXT                     NOT NULL,
    Comments    VARCHAR(255),
    PRIMARY KEY (HymnName, Version, Number)
);