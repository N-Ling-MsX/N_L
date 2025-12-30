<script>
import { bitable } from '@lark-base-open/js-sdk';
import { ref, onMounted, watch } from 'vue';
import {
  ElButton,
  ElForm,
  ElFormItem,
  ElSelect,
  ElOption,
  ElCheckboxGroup,
  ElCheckbox,
  ElDivider,
  ElInputNumber,
  ElTabs,
  ElTabPane
} from 'element-plus';

export default {
  components: {
    ElButton,
    ElForm,
    ElFormItem,
    ElSelect,
    ElOption,
    ElCheckboxGroup,
    ElCheckbox,
    ElDivider,
    ElInputNumber,
    ElTabs,
    ElTabPane,
  },
  setup() {
    const formData = ref({
      table: '',
      view: '',
      fields: [],
      template: 'receipt',
      printCount: 10,
      printAll: false,
      batchPrint: false
    });
    const tableMetaList = ref([]);
    const viewList = ref([]);
    const fieldList = ref([]);
    const recordList = ref([]);
    const selectedRecords = ref([]);
    // 用于记录选择的变量
    const selectedRecordNumbers = ref('');
    const previewRecords = ref([]);
    // Tab切换状态变量
    const activeTab = ref('selectPrint');
    const templateList = ref([
      { value: 'receipt', label: '小票机模板' },
      { value: 'label', label: '标签打印模板' },
      { value: 'standard', label: '标准打印模板' },
      { value: 'simple', label: '简洁打印模板' },
      { value: 'detailed', label: '详细打印模板' }
    ]);

    // 获取视图列表
    const getViews = async (tableId) => {
      if (!tableId) {
        viewList.value = [];
        fieldList.value = [];
        return;
      }
      try {
        const table = await bitable.base.getTableById(tableId);
        const views = await table.getViewMetaList();
        viewList.value = views;
        if (views.length > 0) {
          formData.value.view = views[0].id;
          // 自动获取第一个视图的字段
          await getFields(tableId, views[0].id);
        } else {
          viewList.value = [];
          fieldList.value = [];
        }
      } catch (error) {
        console.error('获取视图列表失败:', error);
        viewList.value = [];
        fieldList.value = [];
      }
    };

    // 获取字段列表
    const getFields = async (tableId, viewId) => {
      if (!tableId || !viewId) {
        fieldList.value = [];
        return;
      }
      try {
        const table = await bitable.base.getTableById(tableId);
        const view = await table.getViewById(viewId);
        const fieldMetas = await view.getFieldMetaList();
        fieldList.value = fieldMetas;
        // 默认不选择任何字段
        formData.value.fields = [];
      } catch (error) {
        console.error('获取字段列表失败:', error);
        fieldList.value = [];
      }
    };

    // 单条记录打印功能
    const printSingleRecord = (record, index, total, template, fields, fieldList, hasRealData) => {
      try {
        // 构建单条记录的打印内容
        let printContent = '';

        // 根据模板类型构建不同格式的打印内容
        switch (template) {
          case 'label':
            // 标签打印模板 - 紧凑布局，适合标签打印机
            printContent += '==============================\n';
            printContent += '记录标签\n';
            printContent += '------------------------------\n';
            if (!hasRealData) {
              printContent += '⚠️  模拟数据\n';
              printContent += '------------------------------\n';
            }
            printContent += `记录 ${index + 1}/${total}\n`;
            printContent += '------------------------------\n';

            fields.forEach(fieldId => {
              const fieldMeta = fieldList.find(f => f.id === fieldId);
              try {
                let value = null;
                // 尝试多种方式获取字段值
                if (typeof record.getCellValue === 'function') {
                  value = record.getCellValue(fieldId);
                } else if (record.data && record.data[fieldId]) {
                  value = record.data[fieldId];
                } else if (record.fields && record.fields[fieldId]) {
                  value = record.fields[fieldId];
                } else {
                  const fieldName = fieldMeta?.name;
                  if (fieldName && record.data && record.data[fieldName]) {
                    value = record.data[fieldName];
                  }
                }
                printContent += `${fieldMeta?.name || fieldId}: ${formatValue(value)}\n`;
              } catch (e) {
                console.error(`获取字段 ${fieldId} 的值失败:`, e);
                if (fieldMeta?.name && fieldMeta.name.includes('负责人') || fieldMeta?.type === 'user') {
                  printContent += `${fieldMeta?.name || fieldId}: [无数据]\n`;
                } else {
                  printContent += `${fieldMeta?.name || fieldId}: [数据获取失败]\n`;
                }
              }
            });
            printContent += '==============================\n\n';
            break;

          case 'standard':
            // 标准打印模板 - 清晰易读的标准格式
            printContent += '\n' + '='.repeat(40) + '\n';
            printContent += '打印记录\n';
            printContent += '='.repeat(40) + '\n';
            if (!hasRealData) {
              printContent += '⚠️  模拟数据测试\n';
              printContent += '-'.repeat(40) + '\n';
            }
            printContent += `\n记录 ${index + 1}/${total}:\n`;
            printContent += '-'.repeat(35) + '\n';

            fields.forEach(fieldId => {
              const fieldMeta = fieldList.find(f => f.id === fieldId);
              try {
                let value = null;
                if (typeof record.getCellValue === 'function') {
                  value = record.getCellValue(fieldId);
                } else if (record.data && record.data[fieldId]) {
                  value = record.data[fieldId];
                } else if (record.fields && record.fields[fieldId]) {
                  value = record.fields[fieldId];
                } else {
                  const fieldName = fieldMeta?.name;
                  if (fieldName && record.data && record.data[fieldName]) {
                    value = record.data[fieldName];
                  }
                }
                // 使用制表符对齐字段名和值
                printContent += `${fieldMeta?.name || fieldId}:\t${formatValue(value)}\n`;
              } catch (e) {
                console.error(`获取字段 ${fieldId} 的值失败:`, e);
                if (fieldMeta?.name && fieldMeta.name.includes('负责人') || fieldMeta?.type === 'user') {
                  printContent += `${fieldMeta?.name || fieldId}:\t[无数据]\n`;
                } else {
                  printContent += `${fieldMeta?.name || fieldId}:\t[数据获取失败]\n`;
                }
              }
            });
            printContent += '\n' + '='.repeat(40) + '\n';
            printContent += `打印时间: ${new Date().toLocaleString()}\n`;
            printContent += '='.repeat(40) + '\n';
            break;

          case 'simple':
            // 简洁打印模板 - 只显示关键信息，无多余装饰
            printContent += `记录 ${index + 1}/${total}:\n`;
            fields.forEach(fieldId => {
              const fieldMeta = fieldList.find(f => f.id === fieldId);
              try {
                let value = null;
                if (typeof record.getCellValue === 'function') {
                  value = record.getCellValue(fieldId);
                } else if (record.data && record.data[fieldId]) {
                  value = record.data[fieldId];
                } else if (record.fields && record.fields[fieldId]) {
                  value = record.fields[fieldId];
                } else {
                  const fieldName = fieldMeta?.name;
                  if (fieldName && record.data && record.data[fieldName]) {
                    value = record.data[fieldName];
                  }
                }
                const formattedValue = formatValue(value);
                // 只显示有值的字段
                if (formattedValue) {
                  printContent += `${fieldMeta?.name || fieldId}: ${formattedValue}\n`;
                }
              } catch (e) {
                // 简洁模式下忽略错误
              }
            });
            printContent += '\n';
            break;

          case 'detailed':
            // 详细打印模板 - 包含所有信息和详细分隔符
            printContent += '\n' + '='.repeat(50) + '\n';
            printContent += '详细打印记录\n';
            printContent += '='.repeat(50) + '\n';
            if (!hasRealData) {
              printContent += '⚠️  模拟数据测试\n';
              printContent += '-'.repeat(50) + '\n';
            }
            printContent += `\n记录 ${index + 1}/${total}\n`;
            printContent += `记录索引: ${index}\n`;
            printContent += `打印时间: ${new Date().toLocaleString()}\n`;
            printContent += '-'.repeat(50) + '\n';

            fields.forEach(fieldId => {
              const fieldMeta = fieldList.find(f => f.id === fieldId);
              printContent += '\n' + '→ ' + (fieldMeta?.name || fieldId) + ':\n';
              try {
                let value = null;
                if (typeof record.getCellValue === 'function') {
                  value = record.getCellValue(fieldId);
                } else if (record.data && record.data[fieldId]) {
                  value = record.data[fieldId];
                } else if (record.fields && record.fields[fieldId]) {
                  value = record.fields[fieldId];
                } else {
                  const fieldName = fieldMeta?.name;
                  if (fieldName && record.data && record.data[fieldName]) {
                    value = record.data[fieldName];
                  }
                }
                // 详细模式下对复杂数据进行更友好的显示
                let formattedValue;
                if (value === null || value === undefined) {
                  formattedValue = '【空】';
                } else if (typeof value === 'object') {
                  // 先检查是否是文本类型的JSON结构
                  if (Array.isArray(value) && value.length > 0 && value[0].type === 'text' && value[0].text) {
                    formattedValue = value[0].text;
                  } else if (!Array.isArray(value) && value.type === 'text' && value.text) {
                    formattedValue = value.text;
                  }
                  // 检查是否是包含value字段的结构（如唯一编号）
                  else if (value.value !== undefined && (typeof value.value === 'string' || typeof value.value === 'number')) {
                    formattedValue = String(value.value);
                  } else if (Array.isArray(value)) {
                    // 处理其他数组类型
                    formattedValue = '[' + value.map(v => {
                      // 对数组中的每个对象也进行文本类型检查
                      if (typeof v === 'object' && v !== null) {
                        if (v.type === 'text' && v.text) return v.text;
                        if (v.value !== undefined) return String(v.value);
                        return JSON.stringify(v);
                      }
                      return v;
                    }).join(', ') + ']';
                  } else {
                    // 处理其他对象类型
                    formattedValue = JSON.stringify(value, null, 2);
                  }
                } else {
                  formattedValue = String(value);
                }
                printContent += '  ' + formattedValue + '\n';
              } catch (e) {
                console.error(`获取字段 ${fieldId} 的值失败:`, e);
                printContent += '  [数据获取失败 - ' + e.message + ']\n';
              }
            });
            printContent += '\n' + '='.repeat(50) + '\n\n';
            break;

          default:
            // 小票机模板（默认）
            printContent += '\n' + '='.repeat(30) + '\n';
            printContent += '打印数据\n';
            printContent += '='.repeat(30) + '\n';

            if (!hasRealData) {
              printContent += '⚠️  模拟数据测试\n';
              printContent += '-'.repeat(30) + '\n';
            }

            printContent += `\n记录 ${index + 1}/${total}:\n`;
            printContent += '-'.repeat(25) + '\n';

            fields.forEach(fieldId => {
              const fieldMeta = fieldList.find(f => f.id === fieldId);
              try {
                let value = null;
                // 尝试多种方式获取字段值
                if (typeof record.getCellValue === 'function') {
                  // 标准方式
                  value = record.getCellValue(fieldId);
                } else if (record.data && record.data[fieldId]) {
                  // 备选方式：直接访问data属性
                  value = record.data[fieldId];
                } else if (record.fields && record.fields[fieldId]) {
                  // 备选方式：直接访问fields属性
                  value = record.fields[fieldId];
                } else {
                  // 如果都失败，尝试使用字段名作为键
                  const fieldName = fieldMeta?.name;
                  if (fieldName && record.data && record.data[fieldName]) {
                    value = record.data[fieldName];
                  }
                }
                printContent += `${fieldMeta?.name || fieldId}: ${formatValue(value)}\n`;
              } catch (e) {
                console.error(`获取字段 ${fieldId} 的值失败:`, e);
                // 为人员类型字段提供特殊处理
                if (fieldMeta?.name && fieldMeta.name.includes('负责人') || fieldMeta?.type === 'user') {
                  printContent += `${fieldMeta?.name || fieldId}: [无数据]\n`;
                } else {
                  printContent += `${fieldMeta?.name || fieldId}: [数据获取失败]\n`;
                }
              }
            });

            printContent += '\n' + '='.repeat(30) + '\n';
            printContent += `打印时间: ${new Date().toLocaleString()}\n`;
            printContent += '='.repeat(30) + '\n';
        }

        // 创建打印窗口
        const printWindow = window.open('', '_blank', 'width=600,height=400');

        if (!printWindow) {
          alert('无法打开打印窗口，请检查浏览器弹窗设置并允许弹窗');
          return false;
        }

        // 构建完整的HTML字符串一次性写入
        const htmlContent = `
        <html>
          <head>
            <title>打印预览</title>
            <style>body{font-family:monospace;font-size:12px;white-space:pre-wrap;padding:20px;}</style>
          </head>
          <body>
            <pre>${printContent}</pre>
            <script>
              window.onload = function() {
                setTimeout(function() {
                  window.print();
                  window.close();
                }, 500);
              };
            <` + `/script>
          </body>
        </html>
      `;

        // 使用innerHTML方式写入内容，更可靠
        printWindow.document.documentElement.innerHTML = htmlContent;

        // 触发打印
        setTimeout(() => {
          try {
            printWindow.print();
            // 打印对话框关闭后自动关闭窗口
            printWindow.onafterprint = function () {
              printWindow.close();
            };
          } catch (printError) {
            console.error('打印执行失败:', printError);
            alert('打印执行失败');
            printWindow.close();
          }
        }, 100);

        return true;
      } catch (error) {
        console.error('单条记录打印失败:', error);
        return false;
      }
    };

    // 打印功能
    const printData = async () => {
      const { table, view, fields, template, batchPrint } = formData.value;
      if (!table || !view || fields.length === 0) {
        alert('请选择数据表、视图和字段');
        return;
      }

      try {
        let records = [];
        let hasRealData = true;

        // 首先检查是否有通过输入框选择的记录（预览记录）
        if (previewRecords.value.length > 0) {
          records = previewRecords.value.map(preview => preview.record);
        } else if (selectedRecords.value.length > 0) {
          // 如果没有预览记录，但有传统方式选中的记录
          records = selectedRecords.value;
        } else {
          // 尝试获取真实数据
          try {
            const tableObj = await bitable.base.getTableById(table);
            const viewObj = await tableObj.getViewById(view);
            const recordIds = await viewObj.getVisibleRecordIdList();
            // 根据选择的打印条数截取记录ID列表
            const { printCount, printAll } = formData.value;
            const targetRecordIds = printAll ? recordIds : recordIds.slice(0, printCount);
            records = await Promise.all(
              targetRecordIds.map(id => tableObj.getRecordById(id))
            );

            if (records.length === 0) {
              hasRealData = false;
            }
          } catch (apiError) {
            console.warn('获取真实数据失败，使用模拟数据:', apiError);
            hasRealData = false;
          }

          // 如果没有真实数据，使用模拟数据
          if (!hasRealData) {
            alert('注意：当前环境无法获取真实数据，将使用模拟数据进行打印测试');

            // 创建模拟记录
            const mockRecords = [];
            const { printCount, printAll } = formData.value;
            const mockRecordCount = printAll ? 10 : (printCount > 0 ? printCount : 5);
            for (let i = 0; i < mockRecordCount; i++) {
              const mockRecord = {
                getCellValue: (fieldId) => {
                  const fieldMeta = fieldList.value.find(f => f.id === fieldId);
                  const fieldName = fieldMeta?.name || '字段';
                  return `${fieldName} - 模拟数据 ${i + 1}`;
                }
              };
              mockRecords.push(mockRecord);
            }
            records = mockRecords;
          }
        }

        // 添加批量打印逻辑
        if (batchPrint) {
          const confirmMessage = `即将批量打印 ${records.length} 条记录，是否继续？`;
          if (!confirm(confirmMessage)) {
            return;
          }

          // 批量打印模式 - 一条一条打印
          alert(`即将开始批量打印 ${records.length} 条记录，请确保浏览器允许弹窗`);

          // 依次打印每条记录，使用setTimeout控制间隔
          records.forEach((record, index) => {
            setTimeout(() => {
              printSingleRecord(record, index, records.length, template, fields, fieldList.value, hasRealData);
            }, index * 2000); // 每条记录间隔2秒
          });

          return;
        }

        // 普通打印模式 - 构建打印内容
        let printContent = '';

        // 根据模板类型构建不同格式的打印内容
        switch (template) {
          case 'label':
            // 标签打印模板 - 紧凑布局
            if (!hasRealData) {
              printContent += '⚠️  模拟数据测试\n';
              printContent += '==============================\n';
            }

            records.forEach((record, index) => {
              printContent += '==============================\n';
              printContent += '记录标签\n';
              printContent += '------------------------------\n';
              printContent += `记录 ${index + 1}/${records.length}\n`;
              printContent += '------------------------------\n';

              fields.forEach(fieldId => {
                const fieldMeta = fieldList.value.find(f => f.id === fieldId);
                try {
                  let value = null;
                  if (typeof record.getCellValue === 'function') {
                    value = record.getCellValue(fieldId);
                  } else if (record.data && record.data[fieldId]) {
                    value = record.data[fieldId];
                  } else if (record.fields && record.fields[fieldId]) {
                    value = record.fields[fieldId];
                  } else {
                    const fieldName = fieldMeta?.name;
                    if (fieldName && record.data && record.data[fieldName]) {
                      value = record.data[fieldName];
                    }
                  }
                  printContent += `${fieldMeta?.name || fieldId}: ${formatValue(value)}\n`;
                } catch (e) {
                  console.error(`获取字段 ${fieldId} 的值失败:`, e);
                  if (fieldMeta?.name && fieldMeta.name.includes('负责人') || fieldMeta?.type === 'user') {
                    printContent += `${fieldMeta?.name || fieldId}: [无数据]\n`;
                  } else {
                    printContent += `${fieldMeta?.name || fieldId}: [数据获取失败]\n`;
                  }
                }
              });
              printContent += '==============================\n\n';
            });
            break;

          case 'standard':
            // 标准打印模板 - 清晰易读
            printContent += '\n' + '='.repeat(40) + '\n';
            printContent += '批量打印记录\n';
            printContent += '='.repeat(40) + '\n';
            if (!hasRealData) {
              printContent += '⚠️  模拟数据测试\n';
              printContent += '-'.repeat(40) + '\n';
            }

            records.forEach((record, index) => {
              printContent += `\n记录 ${index + 1}/${records.length}:\n`;
              printContent += '-'.repeat(35) + '\n';

              fields.forEach(fieldId => {
                const fieldMeta = fieldList.value.find(f => f.id === fieldId);
                try {
                  let value = null;
                  if (typeof record.getCellValue === 'function') {
                    value = record.getCellValue(fieldId);
                  } else if (record.data && record.data[fieldId]) {
                    value = record.data[fieldId];
                  } else if (record.fields && record.fields[fieldId]) {
                    value = record.fields[fieldId];
                  } else {
                    const fieldName = fieldMeta?.name;
                    if (fieldName && record.data && record.data[fieldName]) {
                      value = record.data[fieldName];
                    }
                  }
                  printContent += `${fieldMeta?.name || fieldId}:\t${formatValue(value)}\n`;
                } catch (e) {
                  console.error(`获取字段 ${fieldId} 的值失败:`, e);
                  if (fieldMeta?.name && fieldMeta.name.includes('负责人') || fieldMeta?.type === 'user') {
                    printContent += `${fieldMeta?.name || fieldId}:\t[无数据]\n`;
                  } else {
                    printContent += `${fieldMeta?.name || fieldId}:\t[数据获取失败]\n`;
                  }
                }
              });
            });
            printContent += '\n' + '='.repeat(40) + '\n';
            printContent += `总记录数: ${records.length}\n`;
            printContent += `打印时间: ${new Date().toLocaleString()}\n`;
            printContent += '='.repeat(40) + '\n';
            break;

          case 'simple':
            // 简洁打印模板 - 只显示关键信息
            printContent += `批量打印记录 (共 ${records.length} 条)\n\n`;
            if (!hasRealData) {
              printContent += '⚠️  模拟数据测试\n\n';
            }

            records.forEach((record, index) => {
              printContent += `记录 ${index + 1}:\n`;
              fields.forEach(fieldId => {
                const fieldMeta = fieldList.value.find(f => f.id === fieldId);
                try {
                  let value = null;
                  if (typeof record.getCellValue === 'function') {
                    value = record.getCellValue(fieldId);
                  } else if (record.data && record.data[fieldId]) {
                    value = record.data[fieldId];
                  } else if (record.fields && record.fields[fieldId]) {
                    value = record.fields[fieldId];
                  } else {
                    const fieldName = fieldMeta?.name;
                    if (fieldName && record.data && record.data[fieldName]) {
                      value = record.data[fieldName];
                    }
                  }
                  const formattedValue = formatValue(value);
                  if (formattedValue) {
                    printContent += `${fieldMeta?.name || fieldId}: ${formattedValue}\n`;
                  }
                } catch (e) {
                  // 简洁模式下忽略错误
                }
              });
              printContent += '\n';
            });
            break;

          case 'detailed':
            // 详细打印模板 - 完整信息
            printContent += '\n' + '='.repeat(50) + '\n';
            printContent += '详细批量打印记录\n';
            printContent += '='.repeat(50) + '\n';
            if (!hasRealData) {
              printContent += '⚠️  模拟数据测试\n';
              printContent += '-'.repeat(50) + '\n';
            }
            printContent += `总记录数: ${records.length}\n`;
            printContent += `打印时间: ${new Date().toLocaleString()}\n`;
            printContent += '-'.repeat(50) + '\n';

            records.forEach((record, index) => {
              printContent += '\n' + '='.repeat(40) + '\n';
              printContent += `记录 ${index + 1}/${records.length}\n`;
              printContent += '-'.repeat(40) + '\n';

              fields.forEach(fieldId => {
                const fieldMeta = fieldList.value.find(f => f.id === fieldId);
                printContent += '\n' + '→ ' + (fieldMeta?.name || fieldId) + ':\n';
                try {
                  let value = null;
                  if (typeof record.getCellValue === 'function') {
                    value = record.getCellValue(fieldId);
                  } else if (record.data && record.data[fieldId]) {
                    value = record.data[fieldId];
                  } else if (record.fields && record.fields[fieldId]) {
                    value = record.fields[fieldId];
                  } else {
                    const fieldName = fieldMeta?.name;
                    if (fieldName && record.data && record.data[fieldName]) {
                      value = record.data[fieldName];
                    }
                  }
                  let formattedValue;
                  if (value === null || value === undefined) {
                    formattedValue = '【空】';
                  } else if (typeof value === 'object') {
                    // 先检查是否是文本类型的JSON结构
                    if (Array.isArray(value) && value.length > 0 && value[0].type === 'text' && value[0].text) {
                      formattedValue = value[0].text;
                    } else if (!Array.isArray(value) && value.type === 'text' && value.text) {
                      formattedValue = value.text;
                    }
                    // 检查是否是包含value字段的结构（如唯一编号）
                    else if (value.value !== undefined && (typeof value.value === 'string' || typeof value.value === 'number')) {
                      formattedValue = String(value.value);
                    } else if (Array.isArray(value)) {
                      // 处理其他数组类型
                      formattedValue = '[' + value.map(v => {
                        // 对数组中的每个对象也进行文本类型检查
                        if (typeof v === 'object' && v !== null) {
                          if (v.type === 'text' && v.text) {
                            return v.text;
                          }
                          if (v.value !== undefined) {
                            return String(v.value);
                          }
                          return JSON.stringify(v);
                        }
                        return v;
                      }).join(', ') + ']';
                    } else {
                      // 处理其他对象类型
                      formattedValue = JSON.stringify(value, null, 2);
                    }
                  } else {
                    formattedValue = String(value);
                  }
                  printContent += '  ' + formattedValue + '\n';
                } catch (e) {
                  console.error(`获取字段 ${fieldId} 的值失败:`, e);
                  printContent += '  [数据获取失败 - ' + e.message + ']\n';
                }
              });
            });
            printContent += '\n' + '='.repeat(50) + '\n';
            printContent += `打印完成\n`;
            printContent += '='.repeat(50) + '\n';
            break;

          default:
            // 小票机模板（默认）
            printContent += '\n' + '='.repeat(30) + '\n';
            printContent += '打印数据\n';
            printContent += '='.repeat(30) + '\n';

            if (!hasRealData) {
              printContent += '⚠️  模拟数据测试\n';
              printContent += '-'.repeat(30) + '\n';
            }

            records.forEach((record, index) => {
              printContent += `\n记录 ${index + 1}:\n`;
              printContent += '-'.repeat(25) + '\n';

              fields.forEach(fieldId => {
                const fieldMeta = fieldList.value.find(f => f.id === fieldId);
                try {
                  let value = null;
                  if (typeof record.getCellValue === 'function') {
                    value = record.getCellValue(fieldId);
                  } else if (record.data && record.data[fieldId]) {
                    value = record.data[fieldId];
                  } else if (record.fields && record.fields[fieldId]) {
                    value = record.fields[fieldId];
                  } else {
                    const fieldName = fieldMeta?.name;
                    if (fieldName && record.data && record.data[fieldName]) {
                      value = record.data[fieldName];
                    }
                  }
                  printContent += `${fieldMeta?.name || fieldId}: ${formatValue(value)}\n`;
                } catch (e) {
                  console.error(`获取字段 ${fieldId} 的值失败:`, e);
                  if (fieldMeta?.name && fieldMeta.name.includes('负责人') || fieldMeta?.type === 'user') {
                    printContent += `${fieldMeta?.name || fieldId}: [无数据]\n`;
                  } else {
                    printContent += `${fieldMeta?.name || fieldId}: [数据获取失败]\n`;
                  }
                }
              });
            });

            printContent += '\n' + '='.repeat(30) + '\n';
            printContent += `打印时间: ${new Date().toLocaleString()}\n`;
            printContent += '='.repeat(30) + '\n';
        }

        // 创建打印窗口 - 使用更可靠的方式
        const printWindow = window.open('', '_blank', 'width=600,height=400,menubar=no,toolbar=no,location=no');

        if (!printWindow) {
          alert('无法打开打印窗口，请检查浏览器弹窗设置并允许弹窗');
          return;
        }

        // 安全地写入内容
        try {
          // 构建完整的HTML字符串一次性写入
          const htmlContent = `
              <html>
                <head>
                  <title>打印预览</title>
                  <style>body{font-family:monospace;font-size:12px;white-space:pre-wrap;padding:20px;}</style>
                </head>
                <body>
                  <pre>${printContent}</pre>
                  <script>
                    window.onload = function() {
                      setTimeout(function() {
                        window.print();
                        window.close();
                      }, 500);
                    };
                  <` + `/script>
                </body>
              </html>
            `;

          // 使用innerHTML方式写入内容，更可靠
          printWindow.document.documentElement.innerHTML = htmlContent;

          // 触发打印
          setTimeout(() => {
            try {
              printWindow.print();
              // 打印对话框关闭后自动关闭窗口
              printWindow.onafterprint = function () {
                printWindow.close();
              };
            } catch (printError) {
              console.error('打印执行失败:', printError);
              alert('打印执行失败');
              printWindow.close();
            }
          }, 100);
        } catch (writeError) {
          console.error('写入打印内容失败:', writeError);
          alert('打印内容生成失败');
          if (printWindow) {
            printWindow.close();
          }
          return;
        }
      } catch (error) {
        console.error('打印失败:', error);
        alert('打印失败: ' + error.message + '\n\n请检查浏览器控制台获取详细信息');
      }
    };

    // 格式化字段值
    const formatValue = (value) => {
      if (value === null || value === undefined) {
        return '';
      }
      if (typeof value === 'object') {
        // 处理人员类型字段
        if (Array.isArray(value) && value.length > 0 && value[0].name) {
          return value.map(user => user.name).join(', ');
        }
        // 处理文本类型的JSON结构（如 {"type":"text","text":"内容"} 或 [{"type":"text","text":"内容"}]）
        if (Array.isArray(value) && value.length > 0 && value[0].type === 'text' && value[0].text) {
          return value[0].text;
        } else if (!Array.isArray(value) && value.type === 'text' && value.text) {
          return value.text;
        }
        // 处理包含value字段的JSON结构（如 {"status":"completed","value":"0255A2"}）
        if (value.value !== undefined && (typeof value.value === 'string' || typeof value.value === 'number')) {
          return String(value.value);
        }
        // 处理其他复杂对象
        try {
          return JSON.stringify(value);
        } catch (e) {
          return '[复杂对象]';
        }
      }
      return String(value);
    };

    // 格式化记录预览信息
    const formatRecordPreview = (record, fieldId) => {
      try {
        let value = null;
        // 尝试多种方式获取字段值
        if (typeof record.getCellValue === 'function') {
          // 标准方式
          value = record.getCellValue(fieldId);
        } else if (record.data && record.data[fieldId]) {
          // 备选方式：直接访问data属性
          value = record.data[fieldId];
        } else if (record.fields && record.fields[fieldId]) {
          // 备选方式：直接访问fields属性
          value = record.fields[fieldId];
        }

        const formatted = formatValue(value);
        // 限制预览长度，避免显示过长
        return formatted.length > 15 ? formatted.substring(0, 15) + '...' : formatted;
      } catch (error) {
        console.error(`获取记录预览信息失败:`, error);
        return '[预览失败]';
      }
    };

    // 解析用户输入的记录序号或范围
    const parseSelectedNumbers = (input) => {
      if (!input) {
        return [];
      }

      const selectedIndices = [];
      const parts = input.split(',').map(part => part.trim());

      parts.forEach(part => {
        // 处理范围格式（如：2-5）
        if (part.includes('-')) {
          const [startStr, endStr] = part.split('-').map(s => s.trim());
          const start = parseInt(startStr, 10);
          const end = parseInt(endStr, 10);

          if (!isNaN(start) && !isNaN(end) && start <= end) {
            for (let i = Math.max(1, start); i <= Math.min(end, recordList.value.length); i++) {
              const index = i - 1; // 转换为数组索引
              if (!selectedIndices.includes(index)) {
                selectedIndices.push(index);
              }
            }
          }
        } else {
          // 处理单个序号（如：3）
          const number = parseInt(part, 10);
          if (!isNaN(number) && number >= 1 && number <= recordList.value.length) {
            const index = number - 1; // 转换为数组索引
            if (!selectedIndices.includes(index)) {
              selectedIndices.push(index);
            }
          }
        }
      });

      return selectedIndices.sort((a, b) => a - b);
    };

    // 更新预览记录
    const updatePreview = () => {
      const selectedIndices = parseSelectedNumbers(selectedRecordNumbers.value);

      // 根据选中的索引获取记录
      previewRecords.value = selectedIndices.map(index => ({
        index,
        record: recordList.value[index]
      }));
    };

    // 监听表变化，更新视图
    watch(() => formData.value.table, async (newTableId) => {
      if (newTableId) {
        await getViews(newTableId);
      }
    });

    // 获取记录列表
    const getRecords = async (tableId, viewId) => {
      if (!tableId || !viewId) {
        recordList.value = [];
        selectedRecords.value = [];
        return;
      }
      try {
        const table = await bitable.base.getTableById(tableId);
        const view = await table.getViewById(viewId);
        const recordIds = await view.getVisibleRecordIdList();
        const records = await Promise.all(
          recordIds.map(id => table.getRecordById(id)) // 获取所有可见记录
        );
        recordList.value = records;
        selectedRecords.value = [];
      } catch (error) {
        console.error('获取记录列表失败:', error);
        recordList.value = [];
        selectedRecords.value = [];
      }
    };

    // 全选处理函数
    const handleSelectAll = (val) => {
      selectedRecords.value = val ? [...recordList.value] : [];
    };

    // 监听视图变化，更新字段和记录
    watch(() => formData.value.view, async (newViewId) => {
      if (newViewId && formData.value.table) {
        await Promise.all([
          getFields(formData.value.table, newViewId),
          getRecords(formData.value.table, newViewId)
        ]);
      }
    });

    onMounted(async () => {
      const [tableList, selection] = await Promise.all([bitable.base.getTableMetaList(), bitable.base.getSelection()]);
      formData.value.table = selection.tableId;
      tableMetaList.value = tableList;

      // 如果有选中的表，加载视图、字段和记录
      if (selection.tableId) {
        await getViews(selection.tableId);
        // 如果有视图，加载记录
        if (formData.value.view) {
          await getRecords(selection.tableId, formData.value.view);
        }
      }
    });

    return {
      formData,
      tableMetaList,
      viewList,
      fieldList,
      recordList,
      selectedRecords,
      templateList,
      printData,
      handleSelectAll,
      formatRecordPreview,
      selectedRecordNumbers,
      previewRecords,
      parseSelectedNumbers,
      updatePreview,
      activeTab
    };
  },
};
</script>

<template>
  <el-form ref="form" class="form" :model="formData" label-position="top">
    <!-- <el-form-item label="开发指南">
      <a
        href="https://bytedance.feishu.cn/docx/HazFdSHH9ofRGKx8424cwzLlnZc"
        target="_blank"
        rel="noopener noreferrer"
      >
        多维表格插件开发指南
      </a>
      、
      <a
        href="https://lark-technologies.larksuite.com/docx/HvCbdSzXNowzMmxWgXsuB2Ngs7d"
        target="_blank"
        rel="noopener noreferrer"
      >
        Base Extensions Guide
      </a>
    </el-form-item>
    <el-form-item label="API">
      <a
        href="https://bytedance.feishu.cn/docx/HjCEd1sPzoVnxIxF3LrcKnepnUf"
        target="_blank"
        rel="noopener noreferrer"
      >
        多维表格插件API
      </a>
      、
      <a
        href="https://lark-technologies.larksuite.com/docx/Y6IcdywRXoTYSOxKwWvuLK09sFe"
        target="_blank"
        rel="noopener noreferrer"
      >
        Base Extensions Front-end API
      </a>
    </el-form-item> -->

    <!-- <el-divider /> -->

    <el-form-item label="选择数据表" size="large">
      <el-select v-model="formData.table" placeholder="请选择数据表" style="width: 100%">
        <el-option v-for="meta in tableMetaList" :key="meta.id" :label="meta.name" :value="meta.id" />
      </el-select>
    </el-form-item>

    <el-form-item label="选择视图" size="large">
      <el-select v-model="formData.view" placeholder="请选择视图" style="width: 100%">
        <el-option v-for="view in viewList" :key="view.id" :label="view.name" :value="view.id" />
      </el-select>
    </el-form-item>

    <el-form-item label="选择字段" size="large">
      <el-select v-model="formData.fields" multiple placeholder="请选择字段" style="width: 100%">
        <el-option v-for="field in fieldList" :key="field.id" :label="field.name" :value="field.id" />
      </el-select>
    </el-form-item>

    <el-form-item label="选择模板" size="large">
      <el-select v-model="formData.template" placeholder="请选择模板" style="width: 100%">
        <el-option v-for="template in templateList" :key="template.value" :label="template.label"
          :value="template.value" />
      </el-select>
    </el-form-item>

    <el-form-item label="批量打印设置" size="large">
      <el-checkbox v-model="formData.batchPrint">
        启用批量打印（一条一条打印）
      </el-checkbox>
    </el-form-item>

    <el-tabs v-model="activeTab" type="card" style="margin-bottom: 20px;">
      <el-tab-pane label="选择打印" name="selectPrint">
        <!-- <el-form-item label="打印设置" size="large">
          <div style="display: flex; align-items: center; gap: 20px;">
            <div style="flex: 1;">
              <el-checkbox v-model="formData.printAll" style="margin-bottom: 5px;">打印全部</el-checkbox>
            </div>
            <div style="flex: 1;">
              <el-input-number
                v-model="formData.printCount"
                placeholder="请输入打印条数"
                :min="1"
                :max="1000"
                style="width: 100%"
                :disabled="formData.printAll"
              />
            </div>
          </div>
        </el-form-item> -->

        <el-form-item label="选择记录" size="large" v-if="recordList.length > 0">
          <br />
          <div style="display: flex; flex-direction: column; gap: 10px;">
            <div style="display: flex; gap: 10px; align-items: center;">
              <el-input v-model="selectedRecordNumbers" placeholder="请输入记录序号" style="flex: 1;"
                @input="parseSelectedNumbers" />
              <el-button type="primary" @click="updatePreview">更新预览</el-button>
            </div>
            <div style="color: #606266; font-size: 14px;">
              提示：当前共有 {{ recordList.length }} 条记录，支持单个序号（如：3）或范围（如：2-5），请输入1-{{ recordList.length }}之间的序号
              <br>
              注意：请尽量可以选择1-50条记录，避免打印时间过长
            </div>
            <!-- 记录预览区域 -->
            <div v-if="previewRecords.length > 0"
              style="border: 1px solid #e4e7ed; border-radius: 4px; padding: 15px; background-color: #f5f7fa;">
              <h4 style="margin-top: 0; margin-bottom: 10px;">预览（共 {{ previewRecords.length }} 条记录）：</h4>
              <div style="max-height: 300px; overflow-y: auto;">
                <div v-for="preview in previewRecords" :key="preview.record.id"
                  style="margin-bottom: 15px; padding-bottom: 10px; border-bottom: 1px dashed #e4e7ed;">
                  <div style="font-weight: bold; margin-bottom: 5px;">记录 {{ preview.index + 1 }}:</div>
                  <div style="display: flex; flex-wrap: wrap; gap: 10px;">
                    <template v-for="field in fieldList.slice(0, 5)" :key="field.id">
                      <div style="color: #606266;">
                        <span style="font-weight: 500;">{{ field.name }}:</span> {{ formatRecordPreview(preview.record,
                          field.id) }}
                      </div>
                    </template>
                    <div v-if="fieldList.length > 5" style="color: #909399;">...</div>
                  </div>
                </div>
              </div>
            </div>
            <div v-else-if="selectedRecordNumbers" style="color: #f56c6c; font-size: 14px;">
              没有找到匹配的记录，请检查输入的序号是否正确
            </div>
          </div>
        </el-form-item>
      </el-tab-pane>
      <el-tab-pane label="批量打印" name="batchPrint">
        <el-form-item label="批量打印说明" size="large">
          <div style="color: #606266; font-size: 14px; line-height: 1.6;">
            <p>1. 批量打印将按照设置的间隔时间依次打印每条记录</p>
            <p>2. 请确保浏览器允许弹窗，否则打印将失败</p>
            <p>3. 建议每次批量打印不超过50条记录，避免打印时间过长</p>
            <p>4. 批量打印时，系统将自动使用选择的字段和模板</p>
          </div>
        </el-form-item>

        <el-form-item label="打印设置" size="large">
          <div style="display: flex; align-items: center; gap: 20px;">
            <div style="flex: 1;">
              <el-checkbox v-model="formData.printAll" style="margin-bottom: 5px;">打印全部</el-checkbox>
            </div>
            <div style="flex: 1;">
              <el-input-number v-model="formData.printCount" placeholder="请输入打印条数" :min="1" :max="1000"
                style="width: 100%" :disabled="formData.printAll" />
            </div>
          </div>
        </el-form-item>
      </el-tab-pane>
    </el-tabs>

    <el-button type="primary" size="large" @click="printData" style="margin-top: 20px;">
      打印
    </el-button>
  </el-form>
</template>

<style scoped>
.form :deep(.el-form-item__label) {
  font-size: 16px;
  color: var(--el-text-color-primary);
  margin-bottom: 0;
}

.form :deep(.el-form-item__content) {
  font-size: 16px;
}
</style>
