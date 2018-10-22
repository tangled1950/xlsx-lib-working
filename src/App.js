import React, { Component } from 'react';
import logo from './logo.svg';
import './App.css';
import XLSX from './needToReplace/xlsx';
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
        ], {header:["Student","Subject", "Grade"]});
        ws['!dataValidation'] =  [
            {sqref: 'B2:B99', type: 'list', values: ['2Maths', '2English', '2History', '2Geography', 'Art', 'Science', 'Computers', 'French']},
            {sqref: 'C2:C99', type: 'decimal', operator: 'between', min:1, max: 10},
        ];

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
