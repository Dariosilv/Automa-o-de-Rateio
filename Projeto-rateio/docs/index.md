# Nome do Projeto

Uma automação de rateio de e-mail .

## Funcionalidades

- Lista de funcionalidades
- Exemplo de uso
- Descrição sobre a aplicação

## Tecnologias Utilizadas

- [React](https://reactjs.org/)
- [XLSX.js](https://github.com/SheetJS/sheetjs) para manipulação de planilhas Excel
- [FileSaver.js](https://github.com/eligrey/FileSaver.js/) para salvar arquivos

  return (
    <div>
      <h1>Automação de Rateio</h1>

      {/* Input para abrir o arquivo Excel */}
      <input type="file" accept=".xlsx" onChange={openExcel} />
      
      <br /><br />

      {/* Inputs para inserir dados */}
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
      <br />

      {/* Botões */}
      <button onClick={insertData}>Incluir/Editar</button>
      <button onClick={calculateTotal}>Calcular Total</button>
      <button onClick={saveToExcel}>Salvar no Excel</button>

      {/* Listagem dos dados */}
      <ul>
        {dataList.map((entry, index) => (
          <li key={index}>
            {entry[0]}, {entry[1]}, {entry[2]}, {entry[3]}, {entry[4]}{' '}
            <button onClick={() => deleteData(index)}>Excluir</button>
          </li>
        ))}
      </ul>
    </div>
  );
}

export default App;
