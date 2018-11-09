import React from 'react';
/* eslint import/no-webpack-loader-syntax: off */
import XLSX from '../loaders/xlsx-loader!xlsx';

function readFileData(file) {
	return new Promise(resolve => {
		const fr = new FileReader();
		fr.onload = e => resolve(e.target.result);
		fr.readAsArrayBuffer(file);
	});
}

export default class XLSXDropZone extends React.Component {
	state = {};
	handleDrag() {
		console.info('drag');
		this.setState({...this.state, dragging: true});
	}
	toggleDrag(e, dragging) {
		e.preventDefault();
		e.stopPropagation();
		this.setState({...this.state, dragging});
	}
	onDragOver(e) {
		this.toggleDrag(e, true);
	}

	onDragEnd(e) {
		this.toggleDrag(e, false);
		const files = e.dataTransfer ? e.dataTransfer.files : e.target.files;
		const file = files.length && files[0];
		if(!file.name.match(/\.xlsx?$/)) return;
		this.setState({...this.state, dragging: false, file});
		readFileData(files[0])
			.then(fileData => {
				const ws = XLSX.read(fileData, {type: 'buffer'});
				if(this.props.onChange) this.props.onChange(ws);
			});
	}

	render() {
		const {state} = this;
		return (
			<div className={this.state.dragging ? 'dropzone-active' : 'dropzone'}
				onDragEnd={e => this.onDragEnd(e)}
				onDrop={e => this.onDragEnd(e)}
				onDragOver={e => this.toggleDrag(e, true)}
			>
				<h2> &nbsp; {' '} {state.file && state.file.name}</h2>
				<p>
					Drag a file here, or
					<label>
						<span style={{color: '#0c0', textDecoration: 'underline', cursor: 'pointer'}}> browse </span>
						<input type='file' style={{display: 'none'}} onChange={e => this.onDragEnd(e)}/>
					</label>
				</p>
			</div>
		);
	}
}
