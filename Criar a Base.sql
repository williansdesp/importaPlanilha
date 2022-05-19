CREATE DATABASE willians

USE willians

CREATE TABLE produto (codItem int, codProduto varchar(50), descricaoInterno varchar(150), descricaoExterno varchar(150),
aplicacao varchar(150), dataCadastro date, marcaProduto varchar(50), categoria varchar(50), custoProducao float,
margemLucro float, precoVenda float, precoBalcao float, ncm bigint, cest bigint, medida varchar(5))


SELECT * FROM produto

DROP TABLE produto

