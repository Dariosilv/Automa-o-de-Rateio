import React, { useState } from 'react';
import * as XLSX from 'xlsx'; // Importar todas as funções da biblioteca XLSX
import { saveAs } from 'file-saver'; // Importar o file-saver para salvar o arquivo Excel

function ExcelManager() {
  const [dataList, setDataList] = useState([]);
  const [nome, setNome] = useState('');
  const [setor, setSetor] = useState('');
  const [centroCusto, setCentroCusto] = useState('');
  const [valor, setValor] = useState('');
  const [email, setEmail] = useState('');

  // Função para carregar dados de um arquivo Excel
  const carregarDados = (e) => {
    const file = e.target.files[0];
    const reader = new FileReader();
    reader.onload = (event) => {
      const workbook = XLSX.read(event.target.result, { type: 'binary' });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const data = XLSX.utils.sheet_to_json(sheet, { header: 1 });
      setDataList(data.slice(1)); // Ignorar a primeira linha (cabeçalho)
    };
    reader.readAsBinaryString(file);
  };

  // Função para inserir novos dados no estado da planilha
  const inserirDados = () => {
    const newData = [nome, setor, centroCusto, parseFloat(valor), email]; // Formatar o valor como número
    setDataList([...dataList, newData]); // Atualizar o estado com os novos dados
    // Limpar os campos após inserir os dados
    setNome('');
    setSetor('');
    setCentroCusto('');
    setValor('');
    setEmail('');
  };

  // Função para salvar os dados em um arquivo Excel
  const salvarNoExcel = () => {
    const headers = [['Nome', 'Setor', 'Centro de Custo', 'Valor', 'E-mail']]; // Cabeçalhos da planilha
    const ws = XLSX.utils.aoa_to_sheet([...headers, ...dataList]); // Combinar cabeçalhos com os dados
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Dados');
    const excelBuffer = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    const file = new Blob([excelBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    saveAs(file, 'dados_atualizados.xlsx');
  };

  return (
    <div>
      <h4>Gerenciamento de Excel</h4>
      <input type="file" accept=".xlsx" onChange={carregarDados} />
      <br /><br />
      <label>Nome:</label>
      <input type="text" value={nome} onChange={(e) => setNome(e.target.value)} />
      <br />
      <label>Setor:</label>
      <input type="text" value={setor} onChange={(e) => setSetor(e.target.value)} />
      <br />
      <label>Centro de Custo:</label>
      <input type="text" value={centroCusto} onChange={(e) => setCentroCusto(e.target.value)} />
      <br />
      <label>Valor:</label>
      <input type="number" value={valor} onChange={(e) => setValor(e.target.value)} />
      <br />
      <label>E-mail:</label>
      <input type="email" value={email} onChange={(e) => setEmail(e.target.value)} />
      <br /><br />
      <button onClick={inserirDados}>Inserir Dados</button>
      <button onClick={salvarNoExcel}>Salvar no Excel</button>

      {/* Renderizar a lista de dados inseridos */}
      <ul>
        {dataList.map((data, index) => (
          <li key={index}>
            {`${data[0]}, ${data[1]}, ${data[2]}, R$${data[3]}, ${data[4]}`}
          </li>
        ))}
      </ul>
    </div>
  );
}

export default ExcelManager;
