import XLSX from 'xlsx';

/* 示例输入数据
  // 列信息 从左向右指定每一列对于的属性字段
  const column = ['name', 'age', 'sex', 'work', 'edu'];
  // 表头信息 与主体信息相同 每一项对于一行 可以有空缺 可通过merges合并单元格
  const header = [
    { name: '名称', age: 'AAA',             work: 'bbb'              },
    {              age: '年龄', sex: '性别', work: '经验', edu: '学历' },
  ];

  // 主体信息 每一项对于一行数据
  const data = [
    { name: '张一', age: '18', sex: '女', work: 2, edu: '本科' },
    { name: '张二', age: '23', sex: '男', work: 5, edu: '博士' },
    { name: '张三', age: '34', sex: '女', work: 2, edu: '大专' },
    { name: '张四', age: '56', sex: '男', work: 6, edu: '本科' },
  ];

  // 合并单元格, 每一项合并的信息代表一个矩形 两个坐标分别描述矩形左上角与右下角的坐标
  const merges = [
    [[0, 0], [0, 1]],
    [[1, 0], [2, 0]],
    [[3, 0], [4, 0]],
  ];
*/

class ExcelExport {
  /**
   * @param {Array} header 表格头部
   * @param {Array} column 表格列
   * @param {Array} body 表格数据
   * @param {String} fileName 表格导出名称
   * @param {Array} merges 单元格合并列表
   */
  export = ({
    header = [],
    column = [],
    body = [],
    merges = [],
    fileName = 'excel',
    sheetName = 'Sheet1',
  }) => {
    const styleCell = this.getBorderStyle;
    const styleHeaderCell = this.getHeaderBorderStyle;
    let row = 1;

    const _headers = header
      .map((v, i) =>
        column.map((key, j) => Object.assign(
          {},
          {
            v: v[key],
            position: String.fromCharCode(65 + j) + (i + row),
          }
        ))
      )
      .reduce((prev, next) => prev.concat(next), [])
      .reduce(
        (prev, next) =>
          Object.assign({}, prev, {
            [next.position]: { v: next.v, s: styleHeaderCell },
          }),
        {}
      );
    row += header.length;
    const _body = body
      .map((v, i) =>
        column.map((key, j) => Object.assign(
          {},
          {
            v: v[key],
            position: String.fromCharCode(65 + j) + (i + row),
          }
        ))
      )
      .reduce((prev, next) => prev.concat(next), [])
      .reduce(
        (prev, next) =>
          Object.assign({}, prev, {
            [next.position]: { v: next.v, s: styleCell },
          }),
        {}
      );
    row += body.length;

    const _merges = this.tableMerges(merges);

    const output = Object.assign({}, _headers, _body);

    const outputPos = Object.keys(output);

    const ref = `${outputPos[0]}:${outputPos[outputPos.length - 1]}`;

    const wb = {
      SheetNames: [sheetName],
      Sheets: {
        [sheetName]: Object.assign({}, output, { '!ref': ref, '!merges': _merges }),
      },
    };

    this.save(wb, `${fileName}.xlsx`);
  }

  // 设置合并表头
  tableMerges = (merges) => {
    const _merges = merges.map((positions) => {
      const start = positions[0];
      const end = positions[1];
      return {
        s: {
          c: start[0],
          r: start[1],
        },
        e: {
          c: end[0],
          r: end[1],
        },
      };
    });
    return [..._merges];
  };

  borderAll = {
    top: {
      style: 'thin',
    },
    bottom: {
      style: 'thin',
    },
    left: {
      style: 'thin',
    },
    right: {
      style: 'thin',
    },
  };

  getBorderStyle = { border: this.borderAll };

  getHeaderBorderStyle = {
    border: this.borderAll,
    font: {
      // sz: 18,
      bold: true,
    },
    alignment: {
      horizontal: 'center',
      vertical: 'center',
    },
  };

  save = (wb, fileName) => {
    const wopts = {
      bookType: 'xlsx',
      bookSST: false,
      type: 'binary',
    };
    const xw = XLSX.write(wb, wopts);
    const obj = new Blob([this.s2ab(xw)], {
      type: '',
    });
    const elem = document.createElement('a');
    elem.download = fileName || '下载';
    elem.href = URL.createObjectURL(obj);
    elem.click();
    setTimeout(() => {
      URL.revokeObjectURL(obj);
    }, 100);
  };

  s2ab = (s) => {
    if (typeof ArrayBuffer !== 'undefined') {
      const buf = new ArrayBuffer(s.length);
      const view = new Uint8Array(buf);
      for (let i = 0; i !== s.length; ++i) view[i] = s.charCodeAt(i) & 0xff;
      return buf;
    }
    const buf = new Array(s.length);
    for (let i = 0; i !== s.length; ++i) buf[i] = s.charCodeAt(i) & 0xff;
    return buf;
  };
}

export default ExcelExport;
