import React, { useState } from 'react';
import * as XLSX from 'xlsx';

function ExcelReader() {
  const [columns, setColumns] = useState([]);
  const [data, setData] = useState([]);
  const [startCode, setStartCode] = useState('A');
  const [endCode, setEndCode] = useState('Z');
  const [maxNum, setMaxNum] = useState(15);
  const [currentRoom, setCurrentRoom] = useState(0);
  const [indexColumn, setIndexColumn] = useState('');

  function allocateSeat() {
    console.log(data);
    console.log(columns);
    const filteredData = data.filter(item =>
      item.room_code >= startCode && item.room_code <= endCode
    );
    const sortedData = filteredData.sort((a, b) => {
      if (a.room_code !== b.room_code) {
        return a.room_code.localeCompare(b.room_code);
      } else {
        // 오름차순 정렬로 변경
        return a[indexColumn] < b[indexColumn] ? -1 : 1;
      }
    });

    let currentSeatNum = 1;
    let roomNum = currentRoom;
    let previousRoomCode = null;

    // 기존 data를 업데이트하면서 room_num과 seat_num을 추가
    const updatedData = data.map(item => {
      if (sortedData.includes(item)) {
        if (item.room_code !== previousRoomCode || currentSeatNum > maxNum) {
          roomNum += 1;
          currentSeatNum = 1;
        }

        const updatedItem = {
          ...item,
          room_num: roomNum,
          seat_num: currentSeatNum,
        };

        previousRoomCode = item.room_code;
        currentSeatNum += 1;

        return updatedItem;
      }
      return item;
    });

    setData(updatedData);
    if (!columns.some(column => column.Header === 'room_num')) {
      setColumns([...columns, { Header: 'room_num', accessor: 'room_num' }]);
    }
    if (!columns.some(column => column.Header === 'seat_num')) {
      setColumns(columns => [...columns, { Header: 'seat_num', accessor: 'seat_num' }]);
    }
  }

  const handleFile = (e) => {
    const file = e.target.files[0];
    const reader = new FileReader();
    reader.onload = (event) => {
      const wb = XLSX.read(event.target.result, { type: 'binary' });
      const sheetName = wb.SheetNames[0];
      const worksheet = wb.Sheets[sheetName];
      const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
      if (jsonData.length > 0) {
        // 빈 컬럼 헤더를 제외하고 headers를 생성합니다.
        const headers = jsonData[0].filter(header => header).map((header, index) => ({
          Header: header,
          accessor: header.toLowerCase().replace(/\s/g, '_'),
        }));

        // data 배열을 생성할 때, headers 배열에 맞춰서 각 row의 데이터를 매핑합니다.
        const data = jsonData.slice(1).map(row => {
          let rowData = {};
          headers.forEach((col, index) => {
            rowData[col.accessor] = row[index] || '';
          });
          return rowData;
        });

        setColumns(headers);
        setData(data);
      }
    };
    reader.readAsBinaryString(file);
  };

  const handleChange = (rowIndex, colAccessor, value) => {
    const newData = data.map((row, index) => {
      if (index === rowIndex) {
        return { ...row, [colAccessor]: value };
      }
      return row;
    });
    setData(newData);
  };

  const exportToExcel = () => {
    if (columns.length && data.length) {
      const ws = XLSX.utils.json_to_sheet(data.map(row => {
        let newRow = {};
        columns.forEach(col => {
          newRow[col.Header] = row[col.accessor];
        });
        return newRow;
      }));

      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "Sheet1");
      XLSX.writeFile(wb, "exportedData.xlsx");
    } else {
      alert("No data available to export");
    }
  };

  return (
    <div>
      <input type="file" accept=".xlsx, .xls" onChange={handleFile} />
      {data.length > 0 && <button onClick={exportToExcel}>Export to Excel</button>}
      <button onClick={allocateSeat}>자리세팅</button>
      <br />
      <span>Index Column: </span>
      <select value={indexColumn} onChange={e => setIndexColumn(e.target.value)}>
        {columns.map((column, index) => (
          <option key={index} value={column.Header}>{column.Header}</option>
        ))}
      </select><br />
      <span>시작 room_code</span>
      <select value={startCode} onChange={e => setStartCode(e.target.value)}>
        {Array.from({ length: 26 }, (_, i) => String.fromCharCode(65 + i)).map(letter => (
          <option key={letter} value={letter}>{letter}</option>
        ))}
      </select><br />
      <span>종료 room_code</span>
      <select value={endCode} onChange={e => setEndCode(e.target.value)}>
        {Array.from({ length: 26 }, (_, i) => String.fromCharCode(65 + i)).map(letter => (
          <option key={letter} value={letter}>{letter}</option>
        ))}
      </select><br />
      <span>각 고사실별 최대 인원</span>
      <input
        type="number"
        value={maxNum}
        onChange={e => setMaxNum(Number(e.target.value))}
        min="1"
      /><br />
      <span>시작 고사실 </span>
      <input
        type="number"
        value={currentRoom}
        onChange={e => setCurrentRoom(Number(e.target.value))}
        min="1"
      />
      
      <div style={{ overflowY: 'auto', maxHeight: '400px', border: '1px solid black' }}>
        <table style={{ borderCollapse: 'collapse', width: '100%', tableLayout: 'fixed' }}>
          <thead style={{ display: 'table', width: '100%', tableLayout: 'fixed' }}>
            <tr>
              {columns.map((column, index) => (
                <th key={index} style={{ border: '1px solid black', minWidth: '300px', position: 'sticky', top: 0, background: 'white', zIndex: 1 }}>
                  {column.Header}
                </th>
              ))}
            </tr>
          </thead>
          <tbody style={{ display: 'table', width: '100%', tableLayout: 'fixed' }}>
            {data.map((row, rowIndex) => (
              <tr key={rowIndex}>
                {columns.map((column, ci) => (
                  <td key={`${rowIndex}-${ci}`} style={{ border: '1px solid black', minWidth: '300px' }}>
                    {row[column.accessor]}
                  </td>
                ))}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>

  );
}

export default ExcelReader;
