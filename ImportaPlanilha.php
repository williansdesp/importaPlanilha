<?php 
ini_set('max_execution_time','-1');
require_once "SimpleXLSX.php";

class ImportaPlanilha{

	// Atributo recebe a instância da conexão PDO
	private $conexao  = null;

     // Atributo recebe uma instância da classe SimpleXLSX
	private $planilha = null;

	// Atributo recebe a quantidade de linhas da planilha
	private $linhas   = null;

	// Atributo recebe a quantidade de colunas da planilha
	private $colunas  = null;

	/*
	 * Método Construtor da classe
	 * @param $path - Caminho e nome da planilha do Excel xlsx
	 * @param $conexao - Instância da conexão PDO
	 */
	public function __construct($path=null, $conexao=null){

		if(!empty($path) && file_exists($path)):
			$this->planilha = new SimpleXLSX($path);
			list($this->colunas, $this->linhas) = $this->planilha->dimension();
		else:
			echo 'Arquivo não encontrado!';
			exit();
		endif;

		if(!empty($conexao)):
			$this->conexao = $conexao;
		else:
			echo 'Conexão não informada!';
			exit();
		endif;

	}

	/*
	 * Método que retorna o valor do atributo $linhas
	 * @return Valor inteiro contendo a quantidade de linhas na planilha
	 */
	public function getQtdeLinhas(){
		return $this->linhas;
	}

	/*
	 * Método que retorna o valor do atributo $colunas
	 * @return Valor inteiro contendo a quantidade de colunas na planilha
	 */
	public function getQtdeColunas(){
		return $this->colunas;
	}


	/*
	 * Método para ler os dados da planilha e inserir no banco de dados
	 * @return Valor Inteiro contendo a quantidade de linhas importadas
	 */
	public function insertDados(){

		try{
			$sql = "INSERT INTO produto (codItem, codProduto, descricaoInterno, 
			descricaoExterno, aplicacao, dataCadastro, marcaProduto, categoria, 
			custoProducao, margemLucro, precoVenda, precoBalcao, ncm, cest, medida)
			VALUES(?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)";
			$stm = $this->conexao->prepare($sql);
			
			$linha = 0;
			foreach($this->planilha->rows() as $chave => $valor):
				if ($chave >= 1):		
					$codItem  = trim($valor[0]);
					$codProduto    = trim($valor[1]);
					$descricaoInterno  = trim($valor[2]);
					$descricaoExterno    = trim($valor[3]);
					$aplicacao  = trim($valor[4]);
					$dataCadastro    = trim($valor[5]);
					$marcaProduto  = trim($valor[6]);
					$categoria    = trim($valor[7]);
					$custoProducao  = trim($valor[8]);
					$margemLucro    = trim($valor[9]);
					$precoVenda  = trim($valor[10]);
					$precoBalcao    = trim($valor[11]);
					$ncm  = trim($valor[12]);
					$cest    = trim($valor[13]);
					$medida    = trim($valor[14]);

					$stm->bindValue(1, $codItem);
					$stm->bindValue(2, $codProduto);
					$stm->bindValue(3, $descricaoInterno);
					$stm->bindValue(4, $descricaoExterno);
					$stm->bindValue(5, $aplicacao);
					$stm->bindValue(6, $dataCadastro);
					$stm->bindValue(7, $marcaProduto);
					$stm->bindValue(8, $categoria);
					$stm->bindValue(9, $custoProducao);
					$stm->bindValue(10, $margemLucro);
					$stm->bindValue(11, $precoVenda);
					$stm->bindValue(12, $precoBalcao);
					$stm->bindValue(13, $ncm);
					$stm->bindValue(14, $cest);
					$stm->bindValue(15, $medida);
					$retorno = $stm->execute();
					
					if($retorno == true) $linha++;
				 endif;
			endforeach;

			return $linha;
		}catch(Exception $erro){
			echo 'Erro: ' . $erro->getMessage();
		}

	}
}