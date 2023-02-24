import excelToTable from './src/excelToTable'

// 为组件提供 install 安装方法，供按需引入
excelToTable.install = Vue => {
    Vue.component(excelToTable.name, excelToTable)
}

// 默认导出组件
export default excelToTable