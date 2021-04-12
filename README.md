# SistemaDeVendas
Sistema de Vendas - Desafio LinearSistemas

<h4 align="center">
  Foi desenvolvido uma aplicação básica desktop para um sistema de Vendas, utilizando a linguagem Vb6.
</h4>

<p align="center">
  <a href="#funcionalidades">Funcionalidades</a>&nbsp;&nbsp;&nbsp;|&nbsp;&nbsp;&nbsp;
  <a href="#EstruturaDB">Estrutura de Banco de dados</a>&nbsp;&nbsp;&nbsp;|&nbsp;&nbsp;&nbsp;
  <a href="#memo-license">Licença</a>
</p>

### Funcionalidades

- Criar uma tela de login;
```bash
 RN1: Não deve permitir logar com um usuário que não esteja cadastrado no banco de dados.
```
- Criar uma tela principal MDI com menus;
```bash
 RN1: Deve-se apresentar no formulário, a data corrente de login e o nome do usuário logado. Sugerimos no rodapé, conforme mockup da tela.
```
- Criar um formulário para cadastro de Cliente;
```bash
 RN1: Não permitir excluir um cliente que esteja referenciado em uma venda. 
 RN2: Não permitir gravar um cliente sem o valor de limite de crédito, ou limite 0.
```
- Criar um formulário para cadastro de Produtos;
```bash
 RN1: Não permitir gravar produtos sem o código, ou com código zerado.
 RN2: Não permitir gravar produtos sem o preço, ou com preço zerado.
```
- Criar um formulário para cadastro de Pedidos e seus itens;
```bash
 RN1: Não permitir gravar pedido para clientes cujo o valor do limite, ultrapasse o do cadastro de cliente.
 RN2: Ao gravar um pedido, deve-se abater o valor total do pedido, do limite de crédito do cliente.
 RN3: Não permitir inserção de produtos com preço de venda zerados.
```

### :heavy_check_mark: Configurações necessárias

Seguem as configurações necessárias para visualizar a aplicação em sua máquina.

-  Necessário registrar a OCX RandControls.ocx na pasta c:\Windows\Syswow64, através do prompt de comando em modo administrador:
```bash
Regsvr32 randcontrols.ocx
```
- Maiores informações de regras, consulte: </br>
https://github.com/acessolinear/analista-desenvolvedor/blob/main/README.md

### <h3 id="EstruturaDB">🎲 Estrutura de banco de dados</h3>
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

