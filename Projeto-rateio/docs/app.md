import React, { useState, useEffect } from 'react';
import XLSX from 'xlsx';
import { saveAs } from 'file-saver';

function App() {
  const [dataList, setDataList] = useState([]);
  const [excelFile, setExcelFile] = useState(null);
  const [nome, setNome] = useState('');
  const [setor, setSetor] = useState('');
  const [centroCusto, setCentroCusto] = useState('');
  const [valor, setValor] = useState('');
  const [email, setEmail] = useState('');

  const abrirExcel = () => {
    const fileInput = document.getElementById('fileInput');
    fileInput.click();
  };

  const carregarDados = (file) => {
    const reader = new FileReader();
    reader.onload = (event) => {
      const workbook = XLSX.read(event.target.result, { type: 'array' });
      const sheet = workbook.Sheets['Sheet1'];
      const data = [];
      for (let row = 2; row <= sheet.data.length; row++) {
        data.push([
          sheet.data[row][0].v,
          sheet.data[row][1].v,
          sheet.data[row][2].v,
          sheet.data[row][3].v,
          sheet.data[row][4].v,
        ]);
      }
      setDataList(data);
    };
    reader.readAsArrayBuffer(file);
  };

  const atualizarListbox = () => {
    const listbox = document.getElementById('listbox');
    listbox.innerHTML = '';
    dataList.forEach((entry, index) => {
      const listItem = document.createElement('li');
      listItem.textContent = `${index + 1}. ${entry[0]}, ${entry[1]}, ${entry[2]}, ${entry[3]}, ${entry[4]}`;
      listbox.appendChild(listItem);
    });
  };

  const inserirDados = () => {
    const newData = [nome, setor, centroCusto, valor, email];
    if (dataList.length === 0 || dataList.find((entry) => entry[0] === nome)) {
      dataList.push(newData);
    } else {
      const index = dataList.findIndex((entry) => entry[0] === nome);
      dataList[index] = newData;
    }
    setDataList(dataList);
    atualizarListbox();
  };

  const calcularTotal = () => {
    const total = dataList.reduce((acc, entry) => acc + parseFloat(entry[3]), 0);
    alert(`O valor total é: ${total.toFixed(2)}`);
  };

  const excluirDados = () => {
    const selectedIndex = document.getElementById('listbox').selectedIndex;
    if (selectedIndex !== -1) {
      dataList.splice(selectedIndex, 1);
      setDataList(dataList);
      atualizarListbox();
    } else {
      alert('Selecione um item para excluir!');
    }
  };

  const salvarNoExcel = () => {
    const workbook = XLSX.createWorkbook();
    const sheet = workbook.addSheet('Sheet1');
    dataList.forEach((entry, row) => {
      sheet.addRow([
        { v: entry[0], t: 's' },
        { v: entry[1], t: 's' },
        { v: entry[2], t: 's' },
        { v: entry[3], t: 'n' },
        { v: entry[4], t: 's' },
      ]);
    });
    const excelBuffer = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
    const excelFile = new Blob([excelBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    saveAs(excelFile, 'dados.xlsx');
  };

  return (
    <div>
      <input type="file" id="fileInput" onChange={(e) => carregarDados(e.target.files[0])} />
      <button onClick={abrirExcel}></button>
      <br />
      <br />
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
      <label>E-mail</label>
      <input type="text" value={email} onChange={(e) => setemail(e.target.value)} />
      <br />
       <h4>Enviar dados</h4>
      <p>Enviar os dados para o arquivo Excel.</p>
      <button onClick={salvarNoExcel}>Enviar</button>
    </div>
  

  