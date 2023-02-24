<template>
    <div class="import-excex-content" v-if="!isEmpty">
      <ve-table
        ref="tableRef"
        class="ve-table"
        :style="{ 'word-break': 'normal' }"
        :scroll-width="1600"
        row-key-field-name="rowKey"
        :cell-span-option="cellSpanOption"
        :max-height="maxHeight"
        :fixed-header="true"
        :columns="columns"
        :table-data="tableData"
        :editOption="editOption"
        border-y
        :row-style-option="rowStyleOption"
        :contextmenu-body-option="isRightKey ? contextmenuBodyOption : false"
      />
      <div class="btns">
        <el-radio-group v-model="activeSheet" size="small">
          <el-radio-button
            v-for="item in Object.keys(excelData)"
            :key="item"
            :label="item"
          />
        </el-radio-group>
      </div>
    </div>
    <div v-else>
      <el-empty :image-size="200"></el-empty>
    </div>
  </template>
  
  <script>
  import _ from "lodash";
  export default {
    name: "ExcelToTable",
    props: {
      // excel文件地址
      excelUrl: {
        type: String,
        default:
          "https://125.46.183.213:8178/storage/gy-tduck/37693cfc748049e45d87b8c7d8b9aacd/62847e20ba60419f8cf0f6dd8c401fc1.",
      },
      // 最高高度
      maxHeight: {
        type: Number,
        default: 850,
      },
      // 是否可以编辑
      isEdit: {
        type: Boolean,
        default: () => {
          return true;
        },
      },
      // 右键是否显示
      isRightKey: {
        type: Boolean,
        default: () => {
          return true;
        },
      },
    },
    data() {
      return {
        columns: [],
        tableData: [], // 表格数据
        // 解析excel的数据
        excelData: {},
        // excel是否有数据
        isEmpty: false,
        // 当前的sheet
        activeSheet: null,
        //   合并单元格方法
        cellSpanOption: {
          bodyCellSpan: this.objectSpanMethod,
        },
        rowStyleOption: {
          clickHighlight: true,
          hoverHighlight: true,
        },
        // 是否可以编辑
        editOption: {
          beforeCellValueChange: ({ row, column, changeValue }) => {},
          afterCellValueChange: ({ row, column, changeValue }) => {},
        },
        // contextmenu body option
        contextmenuBodyOption: {
          /*
                      before contextmenu show.
                      In this function,You can change the `contextmenu` options
                      */
          beforeShow: ({
            isWholeRowSelection,
            selectionRangeKeys,
            selectionRangeIndexes,
          }) => {
            console.log("---contextmenu body beforeShow--");
            console.log("isWholeRowSelection::", isWholeRowSelection);
            console.log("selectionRangeKeys::", selectionRangeKeys);
            console.log("selectionRangeIndexes::", selectionRangeIndexes);
          },
          // after menu click
          afterMenuClick: ({
            type,
            selectionRangeKeys,
            selectionRangeIndexes,
          }) => {
            console.log("---contextmenu body afterMenuClick--");
            console.log("type::", type);
            console.log("selectionRangeKeys::", selectionRangeKeys);
            console.log("selectionRangeIndexes::", selectionRangeIndexes);
          },
  
          // contextmenus
          contextmenus: [
            {
              type: "CUT",
            },
            {
              type: "COPY",
            },
            {
              type: "SEPARATOR",
            },
            {
              type: "INSERT_ROW_ABOVE",
            },
            {
              type: "INSERT_ROW_BELOW",
            },
            {
              type: "SEPARATOR",
            },
            {
              type: "REMOVE_ROW",
            },
            {
              type: "EMPTY_ROW",
            },
            {
              type: "EMPTY_CELL",
            },
          ],
        },
      };
    },
    mounted() {
      if (_.isEmpty(this.excelUrl)) {
        throw new Error("excelToTable 组件必须要传入 excelUrl");
      }
      this.getExcelData();
    },
    watch: {
      activeSheet: {
        handler(val) {
          if (!_.isEmpty(val)) {
            const { columns, rows, merges } = this.excelData[val];
            this.tableData = rows;
            this.merges = merges;
            this.columns = columns;
          }
        },
        deep: true,
      },
    },
    created() {},
    methods: {
      //解析excel数据 并返回tableData
      async getExcelData() {
        // 获取从URL解析出来的数据
        const buffer = await this.parseUrl(this.excelUrl);
        // table数据
        let excelData = {};
        const that = this;
        const ExcelJS = require("exceljs");
        const workbook = new ExcelJS.Workbook();
        const { _worksheets } = await workbook.xlsx.load(buffer);
        // 循环页签
        _worksheets.forEach((sheet) => {
          const sheetData = {};
          sheetData.sheetName = sheet.name;
          sheetData.merges = sheet._merges;
          sheetData.rows = [];
          let maxRow = 0;
          sheet.eachRow(function (row, rowNumber) {
            if (rowNumber == 0 || maxRow < row.cellCount) {
              maxRow = row.cellCount;
            }
          });
          sheetData.columns = [
            {
              columnIndex: 0,
              field: "-A",
              key: "-A",
              title: "#",
              edit: true,
              align: "center",
              width: "2%",
              renderBodyCell: ({ row, column, rowIndex, columnIndex }, h) => {
                return rowIndex + 1;
              },
            },
          ];
          for (let index = 0; index < maxRow; index++) {
            sheetData.columns.push({
              columnIndex: index + 1,
              field: String.fromCharCode(65 + index),
              title: String.fromCharCode(65 + index),
              key: String.fromCharCode(65 + index),
              edit: that.isEdit ? true : false,
              type: "INPUT",
              width:
                sheet.getColumn(String.fromCharCode(65 + index))?.width || 8.2,
            });
          }
          sheet.eachRow({ includeEmpty: true }, function (row, rowNumber) {
            const rowData = {};
            row.eachCell({ includeEmpty: true }, function (cell, colNumber) {
              let value = " ";
              if (cell.type == 6) {
                value = cell.result;
              } else {
                value = cell.value;
              }
  
              if (_.isObject(value)) {
                value = cell?.text;
              }
  
              rowData[that.deleteNum(cell._address)] = value;
            });
            sheetData.rows.push(rowData);
          });
          sheetData.rows = sheetData.rows.map((row, index) => ({
            ...row,
            rowKey: index,
          }));
          excelData[sheet.name] = sheetData;
        });
        if (_.isEmpty(excelData)) {
          this.isEmpty = true;
        } else {
          this.excelData = excelData;
          this.activeSheet = Object.keys(excelData)[0];
        }
      },
      // 解析文件
      parseUrl(url) {
        return new Promise((resolve) => {
          let xhr = new XMLHttpRequest();
          xhr.open("get", url, true);
          xhr.responseType = "arraybuffer";
          xhr.onload = function (e) {
            if (xhr.status == 200) {
              resolve(new Uint8Array(xhr.response));
            }
          };
          xhr.send();
        });
      },
      deleteNum(str) {
        let reg = /[0-9]+/g;
        let str1 = str.replace(reg, "");
        return str1;
      },
      // 合并单元格
      objectSpanMethod({ row, column }) {
        const rowIndex = +row.rowKey + 1;
        const columnIndex = +column.columnIndex;
        let flag = false;
        let deleteFlag = false;
        let rowspan = null;
        let colspan = null;
        Object.values(this.merges).forEach((element) => {
          const { left, right, top, bottom } = element.model;
          if (rowIndex == top && columnIndex == left) {
            rowspan = bottom - top + 1;
            colspan = right - left + 1;
            flag = true;
          }
          if (
            top <= rowIndex &&
            rowIndex <= bottom &&
            left <= columnIndex &&
            columnIndex <= right
          ) {
            deleteFlag = true;
          }
        });
        if (flag) {
          return {
            rowspan,
            colspan,
          };
        }
        if (deleteFlag) {
          return [0, 0];
        }
      },
  
      // 导出excel文件
      exportExcel() {
        // 创建excel 设置信息
        const ExcelJS = require("exceljs");
        const workbook = new ExcelJS.Workbook();
        workbook.creator = "YZX";
        workbook.created = new Date(1985, 8, 30);
        // Object.keys(that.excelData).forEach((key) => {
        //   const { rows, merges, columns } = this.excelData[key];
        //   const worksheet = workbook.addWorksheet(key);
        //   rows.forEach((element, index) => {
        //     const copyElement = _.cloneDeep(element);
        //     delete copyElement.id;
        //     delete copyElement["-A"];
        //     // 找到最多的数据的一行
        //     let maxLength = 0;
        //     // 最多数据行的索引值  为列渲染做准备
        //     let maxIndex = 0;
        //     // 去掉数据中的 rowKey
        //     let copySheetData = [];
        //     copySheetData = _.cloneDeep(copyElement).map((item, index) => {
        //       delete item.rowKey;
        //       if (!index || maxLength < Object.keys(item).length) {
        //         maxLength = Object.keys(item).length;
        //         maxIndex = index;
        //       }
        //       return item;
        //     });
        //     // excel 每一个格子的集合 cell
        //     const columns = [];
        //     const header = Object.keys(copySheetData[maxIndex]);
        //     templateData._columns.forEach((item, index) => {
        //       columns.push({
        //         header: undefined,
        //         key: header[index],
        //         width: item.width,
        //       });
        //     });
        //     worksheet.columns = columns;
        //     // 把当前sheet勾选的数据放到sheet里
        //     worksheet.addRows(copySheetData);
        //     // 设置sheet里的样式和数据
        //     worksheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
        //       const rowStyle = templateData._rows[rowNumber - 1];
        //       if (selectedRowKey + 1 >= rowNumber) {
        //         row.height = rowStyle.height;
        //         row.eachCell({ includeEmpty: true }, function (cell, colNumber) {
        //           const cellStyle = rowStyle._cells.find(
        //             (c) => c?._address == cell?._address
        //           );
        //           if (cellStyle) {
        //             cell.style = cellStyle.style;
        //           }
        //         });
        //       } else {
        //         const r = worksheet.getRow(selectedRowKey + 1);
        //         row.eachCell({ includeEmpty: true }, function (cell, colNumber) {
        //           if (!_.isEmpty(cell._value.model.value)) {
        //             cell.style.border = r._cells[colNumber - 1].style.border;
        //           }
        //         });
        //       }
        //     });
        //     // 合并单元格处理
        //     sheetData1.forEach((item, index) => {
        //       // 排查选择的这一行有没有在起点
        //       const filterRowMerges = Object.keys(mergesData).filter((k) => {
        //         const key = k.replace(/[^\d]/g, "");
        //         return key == item.rowKey + 1;
        //       });
        //       if (!_.isEmpty(filterRowMerges)) {
        //         filterRowMerges.forEach((rowMerges) => {
        //           let { top, bottom, left, right } = mergesData[rowMerges]?.model;
        //           const start = String.fromCharCode(64 + left) + (index + 1);
        //           const end =
        //             String.fromCharCode(64 + right) + (bottom - top + index + 1);
        //           worksheet.mergeCells(`${start}:${end}`);
        //         });
        //       }
        //     });
  
        //     // 将准备好的数据生成excel 并下载
        //     workbook.xlsx.writeBuffer().then((data) => {
        //       let blob = new Blob([data], {
        //         type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        //       });
        //       const a = document.createElement("a");
        //       a.href = URL.createObjectURL(blob);
        //       a.download = that.reportFileName + ".xlsx";
        //       document.body.appendChild(a);
        //       a.click();
        //       document.body.removeChild(a);
        //       window.URL.revokeObjectURL(a.href);
        //       that.loading = false;
        //       that.$message({
        //         type: "success",
        //         message: "导出成功",
        //       });
        //     });
        //   });
        // });
      },
    },
  };
  </script>
  
  <style scoped>
  .import-excex-content {
    display: flex;
    flex-direction: column;
    height: 100%;
  }
  .ve-table {
    flex: 1;
  }
  .btns {
    display: flex;
    justify-content: start;
    position: relative;
  }
  </style>
  