import React, { Component } from 'react';
import logo from './logo.svg';
import './App.css';
/* eslint import/no-webpack-loader-syntax: off */
import XLSX from '../loaders/xlsx-loader!xlsx';
//import XLSX from 'data-validation-xlsx';
import saveAs from 'file-saver';


class App extends Component {
  constructor(props){
    super(props);
    this.convert = this.convert.bind(this);
  }
  s2ab(s){
    var buf = new ArrayBuffer(s.length);
    var view = new Uint8Array(buf);
    for(var i = 0 ; i < s.length ; i++) view[i] = s.charCodeAt(i) & 0xFF;
    return buf;
  }
  convert = () => {
    var wb = XLSX.utils.book_new();
    wb.Props = {
      Title : "Excel Dropdown Menu",
      Subject: "Data Validation",
      Author: "Alexandru Faina",
      CreatedDate: new Date()
    };
    wb.SheetNames.push('New Sheet');
    var ws = XLSX.utils.json_to_sheet([
      { Student: "Euan"},
      { Student: "Mary"},
      { Student: "Holly"},
    ], {header:["Student","Subject", "Grade", "Email", "Phone Number", "Date", "Gender"]});
    ws['!dataValidation'] =  [
      {sqref: 'A2:A99', type: 'alphabet'},
      {sqref: 'B2:B99', type: 'list', values: ['Maths', 'English', 'History', 'Geography', 'Art', 'Science', 'Computers', 'French']},
      {sqref: 'C2:C99', type: 'decimal', operator: 'between', min:1, max: 10},
      {sqref: 'D2:D99', type: 'email'},
      {sqref: 'E2:E99', type: 'phone_number'},
      {sqref: 'F2:F99', type: 'date', operator: 'between', start: '1/1/1900', end: '12/31/3000'},
      {sqref: 'G2:G99', type: 'list', values: ['Male', 'Female']}
    ];
    ws['!cols'] = [{wch:16},{wch:16},{wch:16},{wch:16},{wch:16},{wch:16},{wch:16}]
    for(let i=0;i<97;i++){
      ws['F'+(2+i)] = {v:'', z: 'mm/dd/yyyy'};
      ws['E'+(2+i)] = {v:'', z: '@'};
    }
    ws['!ref']="A1:G99";

    wb.Sheets['New Sheet'] = ws;

    var wbout = XLSX.write(wb, {bookType: 'xlsx', type:'binary'});

    saveAs(new Blob([this.s2ab(wbout)], {type: "application/octet-stream"}), 'dropdown.xlsx');
  }
  render() {
    return (
      <div className="App">
        <header className="App-header">
          <img src={logo} className="App-logo" alt="logo" />
          <h1 className="App-title">Welcome to React</h1>
        </header>
        <button onClick={this.convert}>Convert</button>
      </div>
    );
  }
}

export default App;
