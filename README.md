# SistemaDeVendas
Sistema de Vendas - Desafio LinearSistemas

<h4 align="center">
  Foi desenvolvido uma aplicação básica desktop para um sistema de Vendas, utilizando a linguagem Vb6.
</h4>

<p align="center">
  <a href="#funcionalidades">Funcionalidades</a>&nbsp;&nbsp;&nbsp;|&nbsp;&nbsp;&nbsp;
  <a href="#beginner-iniciando-a-aplicação">Iniciando a aplicação</a>&nbsp;&nbsp;&nbsp;|&nbsp;&nbsp;&nbsp;
  <a href="#estruturaBd">Estrutura de Banco de dados</a>&nbsp;&nbsp;&nbsp;|&nbsp;&nbsp;&nbsp;
  <a href="#memo-license">Licença</a>
</p>

### Funcionalidades

- Criar uma tela de login;
- Criar uma tela principal MDI com menus;
- Criar um formulário para cadastro de Cliente;
- Criar um formulário para cadastro de Produtos;
- Criar um formulário para cadastro de Pedidos e seus itens;

### :heavy_check_mark: Configurações necessárias

Seguem as configurações necessárias para visualizar a aplicação em sua máquina.

-  Necessário registrar a OCX RandControls.ocx na pasta c:\Windows\Syswow64, através do prompt de comando em modo administrador:
```bash
Regsvr32 randcontrls.ocx
```

### :beginner: Iniciando a aplicação
1. Abra a aplicação no ambiente de desenvolvimento VB6.
```bash
# Após carregar corretamente, pressione F5
```
2. Na tela de login informe.
```bash
# Usuario: admin
Senha: $enhaAdmin
```
3. Maiores informações de regras, consulte: </br>
https://github.com/acessolinear/analista-desenvolvedor/blob/main/README.md

### 🎲 Estrutura de banco de dados
1. Foi Utilizado o MySQL 5.7 e ODBC SQL 8.0 32Bits
2. Nome do Squema: VendasLinear

Cadastro de Pessoa
```bash
CREATE TABLE `pessoa` (
	`Codigo` INT(11) NOT NULL,
	`Nome` VARCHAR(100) NULL DEFAULT NULL COLLATE 'latin1_swedish_ci',
	`Telefone` VARCHAR(15) NULL DEFAULT NULL COLLATE 'latin1_swedish_ci',
	`Celular` VARCHAR(15) NULL DEFAULT NULL COLLATE 'latin1_swedish_ci',
	`Tipo` CHAR(1) NULL DEFAULT NULL COLLATE 'latin1_swedish_ci',
	PRIMARY KEY (`Codigo`) USING BTREE
)
COMMENT='Cadastro de Pessoas'
COLLATE='latin1_swedish_ci'
ENGINE=InnoDB
;
```
Cadastro de Cliente
```bash
CREATE TABLE `cliente` (
	`Codigo` INT(11) NOT NULL,
	`CodigoPessoa` INT(11) NULL DEFAULT NULL,
	`LimiteCredito` DOUBLE NULL DEFAULT NULL,
	PRIMARY KEY (`Codigo`) USING BTREE,
	INDEX `FK_Ciente_Pessoa` (`CodigoPessoa`) USING BTREE,
	CONSTRAINT `FK_Ciente_Pessoa` FOREIGN KEY (`CodigoPessoa`) REFERENCES `vendaslinear`.`pessoa` (`Codigo`) ON UPDATE RESTRICT ON DELETE RESTRICT
)
COMMENT='Cadastro de Clientes'
COLLATE='latin1_swedish_ci'
ENGINE=InnoDB
;
```
Cadastro de Usuario
```bash
CREATE TABLE `usuario` (
	`Codigo` INT(11) NOT NULL,
	`Login` VARCHAR(100) NULL DEFAULT NULL COLLATE 'utf8_general_ci',
	`Senha` VARCHAR(50) NULL DEFAULT NULL COLLATE 'utf8_general_ci',
	`CodigoPessoa` INT(11) NULL DEFAULT NULL,
	PRIMARY KEY (`Codigo`) USING BTREE,
	INDEX `FK_Usuario_Pessoa` (`CodigoPessoa`) USING BTREE,
	CONSTRAINT `FK_Usuario_Pessoa` FOREIGN KEY (`CodigoPessoa`) REFERENCES `vendaslinear`.`pessoa` (`Codigo`) ON UPDATE RESTRICT ON DELETE RESTRICT
)
COMMENT='Cadastro de Usuarios'
COLLATE='latin1_swedish_ci'
ENGINE=InnoDB
;
```
Cadastro de Produto
```bash
CREATE TABLE `produto` (
	`Codigo` INT(11) NOT NULL,
	`Nome` VARCHAR(100) NULL DEFAULT NULL COLLATE 'latin1_swedish_ci',
	`Preco` DOUBLE NULL DEFAULT NULL,
	`CodigoExterno` VARCHAR(50) NULL DEFAULT NULL COLLATE 'latin1_swedish_ci',
	PRIMARY KEY (`Codigo`) USING BTREE
)
COMMENT='Cadastro de Produtos'
COLLATE='latin1_swedish_ci'
ENGINE=InnoDB
;
```
Cadastro de Pedido
```bash
CREATE TABLE `pedido` (
	`Codigo` INT(11) NOT NULL,
	`CodigoCliente` INT(11) NULL DEFAULT NULL,
	`ValorTotal` DOUBLE NULL DEFAULT NULL,
	PRIMARY KEY (`Codigo`) USING BTREE,
	INDEX `FK_Produto_Cliente` (`CodigoCliente`) USING BTREE,
	CONSTRAINT `FK_Produto_Cliente` FOREIGN KEY (`CodigoCliente`) REFERENCES `vendaslinear`.`cliente` (`Codigo`) ON UPDATE RESTRICT ON DELETE RESTRICT
)
COMMENT='Cadastro de Pedido de compras de clientes'
COLLATE='latin1_swedish_ci'
ENGINE=InnoDB
;
```
Cadastro de Item de Pedido
```bash
CREATE TABLE `pedidoitem` (
	`CodigoPedido` INT(11) NOT NULL,
	`CodigoProduto` INT(11) NOT NULL,
	`Quantidade` DOUBLE NULL DEFAULT NULL,
	`Preco` DOUBLE NULL DEFAULT NULL,
	`ValorTotal` DOUBLE NULL DEFAULT NULL,
	PRIMARY KEY (`CodigoPedido`, `CodigoProduto`) USING BTREE,
	INDEX `FK_ItemPedido_Produto` (`CodigoProduto`) USING BTREE,
	CONSTRAINT `FK_ItemPedido_Produto` FOREIGN KEY (`CodigoProduto`) REFERENCES `vendaslinear`.`produto` (`Codigo`) ON UPDATE RESTRICT ON DELETE RESTRICT
)
COMMENT='Itens do Pedido'
COLLATE='latin1_swedish_ci'
ENGINE=InnoDB
;
```

### :memo: License
Esse projeto está liberado para uso e alterações.


Feito por Randerson Maurilio 🖤 Contato: https://www.linkedin.com/in/randerson-maur%C3%ADlio-b8053522/

