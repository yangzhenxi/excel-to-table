import Vue from 'vue'
import App from './App.vue'
import { VeTable,  VeLocale} from "vue-easytable";
import zhCN from "vue-easytable/libs/locale/lang/zh-CN";
// import excelToTable from "../components/index";
import "vue-easytable/libs/theme-default/index.css";
import Element from 'element-ui'
import 'element-ui/lib/theme-chalk/index.css';
import excelToTable from "excel-to-table";

Vue.config.productionTip = false
VeLocale.use(zhCN)
Vue.use(VeTable);
Vue.use(Element)
Vue.use(excelToTable)

new Vue({
  render: h => h(App),
}).$mount('#app')













