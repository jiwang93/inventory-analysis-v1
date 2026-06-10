/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useCallback } from 'react';
import * as XLSX from 'xlsx';
import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';
import { 
  Upload, 
  Download, 
  FileText, 
  Table as TableIcon, 
  ArrowRight, 
  CheckCircle2, 
  AlertCircle,
  Hash,
  MapPin,
  Box,
  Package,
  History,
  BarChart3,
  Loader2
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import { jsPDF } from 'jspdf';
import { cn } from './lib/utils';

interface InventoryRow {
  '箱号': string | number;
  '库位名称': string;
  '功率档'?: string | number;
  '工单号'?: string | number;
  '销售单号'?: string | number;
  '箱等级'?: string;
  '箱条码数'?: number;
}

interface AnalysisResult {
  '库位名称': string;
  '箱号数量': number;
  '可入库数量': number;
  '客户分布'?: Record<string, number>;
}

interface MixedPowerAlert {
  location: string;
  powerDetails: Record<string, string[]>; 
}

interface MixedCustomerAlert {
  location: string;
  customerDetails: Record<string, string[]>; // CustomerName -> List of BoxNumbers
}

interface CustomerStat {
  customerName: string;
  boxCount: number;
  locations: string[];
}

export default function App() {
  const STANDARD_CAPACITY = 18;
  const [data, setData] = useState<InventoryRow[]>([]);
  const [results, setResults] = useState<AnalysisResult[]>([]);
  const [mixedPowerAlerts, setMixedPowerAlerts] = useState<MixedPowerAlert[]>([]);
  const [mixedCustomerAlerts, setMixedCustomerAlerts] = useState<MixedCustomerAlert[]>([]);
  const [orderCustomerMap, setOrderCustomerMap] = useState<Record<string, string>>({});
  const [mappingFileName, setMappingFileName] = useState('');
  const [customerStats, setCustomerStats] = useState<CustomerStat[]>([]);
  const [activeTab, setActiveTab] = useState<'location' | 'customer'>('location');
  const [isProcessing, setIsProcessing] = useState(false);
  const [fileName, setFileName] = useState('');
  const [error, setError] = useState<string | null>(null);

  const processFile = (file: File) => {
    setIsProcessing(true);
    setError(null);
    setFileName(file.name);

    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const bstr = e.target?.result;
        const workbook = XLSX.read(bstr, { type: 'binary' });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        
        // 采用兼具精确索引和模糊匹配的高鲁棒性解析方式
        const arrayRows = XLSX.utils.sheet_to_json(worksheet, { header: 1 }) as any[][];

        if (!arrayRows || arrayRows.length < 2) {
          throw new Error('表格中没有足够的数据行');
        }

        const headers = (arrayRows[0] || []).map(h => String(h || '').trim());
        
        // 查找 [箱号] 和 [库位名称]
        let boxIdx = headers.findIndex(h => h === '箱号');
        let locIdx = headers.findIndex(h => h === '库位名称');
        
        // 常用标题模糊检测以增强容错
        if (boxIdx === -1) boxIdx = headers.findIndex(h => h.includes('箱号') || h.includes('箱码'));
        if (boxIdx === -1) boxIdx = headers.findIndex(h => h === '箱');
        if (locIdx === -1) locIdx = headers.findIndex(h => h.includes('库位') || h.includes('仓位'));
        
        if (boxIdx === -1) {
          throw new Error('未在主库存表格首行中找到 [箱号] 列');
        }
        if (locIdx === -1) {
          throw new Error('未在主库存表格首行中找到 [库位名称] 列');
        }

        // 工单号标题的模糊位置 (作为 column J 越界或为空时的 fallback)
        let orderNoFallbackIdx = headers.findIndex(h => h.includes('工单') || h.includes('销售单') || h.toLowerCase().includes('order'));

        const parsedData: InventoryRow[] = [];
        for (let i = 1; i < arrayRows.length; i++) {
          const rowArr = arrayRows[i];
          if (!rowArr || rowArr.length === 0) continue;
          
          // 排除全是空单元格的虚假行
          const isAllEmpty = rowArr.every(val => val === undefined || val === null || String(val).trim() === '');
          if (isAllEmpty) continue;
          
          const boxVal = rowArr[boxIdx] !== undefined && rowArr[boxIdx] !== null ? String(rowArr[boxIdx]).trim() : '';
          const locVal = rowArr[locIdx] !== undefined && rowArr[locIdx] !== null ? String(rowArr[locIdx]).trim() : '';
          
          // 按用户要求：主库存表中每一箱的工单号在 J 列（Column J 是第 10 列，0-based 索引为 9）
          let orderNoVal = '';
          if (rowArr.length > 9 && rowArr[9] !== undefined && rowArr[9] !== null && String(rowArr[9]).trim() !== '') {
            orderNoVal = String(rowArr[9]).trim();
          } else if (orderNoFallbackIdx !== -1 && rowArr[orderNoFallbackIdx] !== undefined && rowArr[orderNoFallbackIdx] !== null) {
            orderNoVal = String(rowArr[orderNoFallbackIdx]).trim();
          }

          let powerIdx = headers.findIndex(h => h.includes('功率'));
          const powerVal = powerIdx !== -1 ? rowArr[powerIdx] : undefined;

          parsedData.push({
            '箱号': boxVal,
            '库位名称': locVal,
            '功率档': powerVal,
            '工单号': orderNoVal,
            '销售单号': orderNoVal
          });
        }

        if (parsedData.length === 0) {
          throw new Error('主库存表中没有成功加载有效数据行');
        }

        setData(parsedData);
        performAnalysis(parsedData, orderCustomerMap);
      } catch (err) {
        setError(err instanceof Error ? err.message : '文件解析失败，请检查格式是否正确');
        console.error(err);
      } finally {
        setIsProcessing(false);
      }
    };
    reader.onerror = () => {
      setError('文件读取失败');
      setIsProcessing(false);
    };
    reader.readAsBinaryString(file);
  };

  const processMappingFile = (file: File) => {
    setError(null);
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const bstr = e.target?.result;
        const workbook = XLSX.read(bstr, { type: 'binary' });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        
        // 规则：“对照表，工单号B列（索引1），对应的F列（索引5），是客户简称”
        const arrayRows = XLSX.utils.sheet_to_json(worksheet, { header: 1 }) as any[][];
        if (!arrayRows || arrayRows.length === 0) {
          throw new Error('对照表为空或格式不正确');
        }

        // 读取首行做模糊寻址 fallback（如果个别单元格由于某些缘故未对齐B/F）
        const firstRow = arrayRows[0] || [];
        const headers = firstRow.map(h => String(h || '').trim());
        
        let orderFallbackIndex = headers.findIndex(h => h.includes('工单') || h.includes('销售单') || h.toLowerCase().includes('order'));
        let customerFallbackIndex = headers.findIndex(h => h.includes('客户') || h.toLowerCase().includes('customer') || h.toLowerCase().includes('client') || h.includes('简称'));

        // 如果首行包含了标志性文字词组，认为它是一行表头，跳过不当作实际数据映射
        let startRowIndex = 0;
        const bCellVal = String(firstRow[1] || '').trim();
        const fCellVal = String(firstRow[5] || '').trim();
        if (bCellVal.includes('单号') || bCellVal.includes('工单') || fCellVal.includes('客户') || fCellVal.includes('简称')) {
          startRowIndex = 1;
        }

        const newMap: Record<string, string> = {};
        for (let i = startRowIndex; i < arrayRows.length; i++) {
          const rowArr = arrayRows[i];
          if (!rowArr || rowArr.length === 0) continue;

          // 用户强规则对应：Column B (index 1) 和 Column F (index 5)
          let orderVal = '';
          let custVal = '';

          if (rowArr.length > 1 && rowArr[1] !== undefined && rowArr[1] !== null) {
            orderVal = String(rowArr[1]).trim();
          }
          if (rowArr.length > 5 && rowArr[5] !== undefined && rowArr[5] !== null) {
            custVal = String(rowArr[5]).trim();
          }

          // 模糊兜底 (如果B和F提取为空且首行探测到了匹配的列字段)
          if (!orderVal && orderFallbackIndex !== -1 && rowArr[orderFallbackIndex] !== undefined && rowArr[orderFallbackIndex] !== null) {
            orderVal = String(rowArr[orderFallbackIndex]).trim();
          }
          if (!custVal && customerFallbackIndex !== -1 && rowArr[customerFallbackIndex] !== undefined && rowArr[customerFallbackIndex] !== null) {
            custVal = String(rowArr[customerFallbackIndex]).trim();
          }

          if (orderVal && custVal) {
            newMap[orderVal] = custVal;
          }
        }

        if (Object.keys(newMap).length === 0) {
          throw new Error('未能在对照表中找到可转换的工单与客户对照数据');
        }

        setOrderCustomerMap(newMap);
        setMappingFileName(file.name);
        
        if (data.length > 0) {
          performAnalysis(data, newMap);
        }
      } catch (err) {
        setError(err instanceof Error ? err.message : '对照表文件解析失败');
        console.error(err);
      }
    };
    reader.readAsBinaryString(file);
  };

  const performAnalysis = (rows: InventoryRow[], currentMap: Record<string, string> = orderCustomerMap) => {
    const locationMap: Record<string, Set<string | number>> = {};
    // 功率检测：库位 -> { 功率档 -> [箱号] }
    const locationPowerData: Record<string, Record<string, Set<string>>> = {};
    // 客户监测：库位 -> { 客户名称 -> [箱号] }
    const locationCustomerData: Record<string, Record<string, Set<string>>> = {};
 
    // 1. 初始化固定的 C-H (1-60) 库位
    const zones = ['C', 'D', 'E', 'F', 'G', 'H'];
    zones.forEach(zone => {
      for (let i = 1; i <= 60; i++) {
        const locName = `${zone}${String(i).padStart(2, '0')}`;
        locationMap[locName] = new Set();
        locationPowerData[locName] = {};
        locationCustomerData[locName] = {};
      }
    });
 
    // 2. 处理数据
    rows.forEach(row => {
      let loc = String(row['库位名称'] || '').trim();
      if (!loc) return;
      
      const match = loc.match(/^([C-H])(\d+)$/i);
      if (match) {
        loc = `${match[1].toUpperCase()}${String(match[2]).padStart(2, '0')}`;
      }
 
      const box = row['箱号'];
      const power = String(row['功率档'] || '未知').trim();
      const orderNo = String(row['工单号'] || row['销售单号'] || '空单号').trim();
      const custName = currentMap[orderNo] || '未映射客户';
      
      if (!locationMap[loc]) {
        locationMap[loc] = new Set();
        locationPowerData[loc] = {};
        locationCustomerData[loc] = {};
      }
 
      if (box !== undefined && box !== null && String(box).trim() !== '') {
        const boxStr = String(box);
        locationMap[loc].add(boxStr);
 
        // 记录库位下的功率分布
        if (!locationPowerData[loc][power]) {
          locationPowerData[loc][power] = new Set();
        }
        locationPowerData[loc][power].add(boxStr);
 
        // 记录库位下的客户分布
        if (!locationCustomerData[loc][custName]) {
          locationCustomerData[loc][custName] = new Set();
        }
        locationCustomerData[loc][custName].add(boxStr);
      }
    });
 
    // 汇总分析结果
    const analysisResults: AnalysisResult[] = Object.entries(locationMap).map(([location, boxSet]) => {
      const boxCount = boxSet.size;
      const customerMapForLoc = locationCustomerData[location] || {};
      const customersDist: Record<string, number> = {};
      Object.entries(customerMapForLoc).forEach(([cust, boxes]) => {
        if (boxes.size > 0) {
          customersDist[cust] = boxes.size;
        }
      });
      return {
        '库位名称': location,
        '箱号数量': boxCount,
        '可入库数量': STANDARD_CAPACITY - boxCount,
        '客户分布': customersDist
      };
    });
 
    // 排序逻辑保持不变...
    analysisResults.sort((a, b) => {
      const nameA = a['库位名称'];
      const nameB = b['库位名称'];
 
      const getRank = (name: string) => {
        const matchZone = name.match(/^([C-H])(\d+)$/i);
        if (matchZone) return 1;
        if (name.toLowerCase().startsWith('wall')) return 2;
        return 3;
      };
 
      const rankA = getRank(nameA);
      const rankB = getRank(nameB);
 
      if (rankA !== rankB) return rankA - rankB;
 
      if (rankA === 1) {
        const matchA = nameA.match(/^([C-H])(\d+)$/i)!;
        const matchB = nameB.match(/^([C-H])(\d+)$/i)!;
        if (matchA[1].toUpperCase() !== matchB[1].toUpperCase()) {
          return matchA[1].toUpperCase().localeCompare(matchB[1].toUpperCase());
        }
        return parseInt(matchA[2], 10) - parseInt(matchB[2], 10);
      }
 
      return nameA.localeCompare(nameB);
    });
 
    // 处理混档警告 (库位内功率档种类 > 1)
    const alerts: MixedPowerAlert[] = [];
    Object.entries(locationPowerData).forEach(([loc, powers]) => {
      const powerGrades = Object.keys(powers);
      if (powerGrades.length > 1) {
        const powerDetails: Record<string, string[]> = {};
        powerGrades.forEach(p => {
          powerDetails[p] = Array.from(powers[p]);
        });
        alerts.push({
          location: loc,
          powerDetails
        });
      }
    });
 
    // 处理客户混载警告
    const customerAlerts: MixedCustomerAlert[] = [];
    Object.entries(locationCustomerData).forEach(([loc, customers]) => {
      const customerNames = Object.keys(customers);
      if (customerNames.length > 1) {
        const customerDetails: Record<string, string[]> = {};
        customerNames.forEach(c => {
          customerDetails[c] = Array.from(customers[c]);
        });
        customerAlerts.push({
          location: loc,
          customerDetails
        });
      }
    });
 
    // 客户箱数统计
    const customerBoxes: Record<string, Set<string>> = {};
    const customerLocs: Record<string, Set<string>> = {};
 
    rows.forEach(row => {
      const box = row['箱号'];
      if (box !== undefined && box !== null && String(box).trim() !== '') {
        const boxStr = String(box);
        const orderNo = String(row['工单号'] || row['销售单号'] || '空单号').trim();
        const custName = currentMap[orderNo] || '未映射客户';
 
        let loc = String(row['库位名称'] || '').trim();
        if (loc) {
          const match = loc.match(/^([C-H])(\d+)$/i);
          if (match) {
            loc = `${match[1].toUpperCase()}${String(match[2]).padStart(2, '0')}`;
          }
        }
 
        if (!customerBoxes[custName]) {
          customerBoxes[custName] = new Set();
        }
        customerBoxes[custName].add(boxStr);
 
        if (loc) {
          if (!customerLocs[custName]) {
            customerLocs[custName] = new Set();
          }
          customerLocs[custName].add(loc);
        }
      }
    });
 
    const stats: CustomerStat[] = Object.entries(customerBoxes).map(([custName, boxSet]) => {
      return {
        customerName: custName,
        boxCount: boxSet.size,
        locations: Array.from(customerLocs[custName] || [])
      };
    }).sort((a, b) => b.boxCount - a.boxCount);
 
    setResults(analysisResults);
    setMixedPowerAlerts(alerts);
    setMixedCustomerAlerts(customerAlerts);
    setCustomerStats(stats);
  };

  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) processFile(file);
  };

  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault();
    const file = e.dataTransfer.files?.[0];
    if (file) processFile(file);
  };

  const handleMappingFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) processMappingFile(file);
  };

  const handleMappingDrop = (e: React.DragEvent) => {
    e.preventDefault();
    const file = e.dataTransfer.files?.[0];
    if (file) processMappingFile(file);
  };

  const clearMapping = () => {
    setOrderCustomerMap({});
    setMappingFileName('');
    setCustomerStats([]);
    if (data.length > 0) {
      performAnalysis(data, {});
    }
  };

  const exportToExcel = async () => {
    if (results.length === 0) return;
    
    // 找出所有存在于结果中的唯一客户名称
    const uniqueCustomersSet = new Set<string>();
    results.forEach((row) => {
      const dist = row['客户分布'];
      if (dist) {
        Object.keys(dist).forEach((cust) => {
          if (cust && cust.trim() !== '') {
            uniqueCustomersSet.add(cust.trim());
          }
        });
      }
    });
    const uniqueCustomers = Array.from(uniqueCustomersSet).sort();

    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('库位统计分析');

    // 1. 定义基础列
    const baseColumns = [
      { header: '库位名称', key: 'location', width: 20 },
      { header: '箱号数量', key: 'count', width: 15 },
      { header: '可入库数量 (标准18)', key: 'available', width: 25 },
    ];

    // 2. 为每个客户动态追加一列
    const customerColumns = uniqueCustomers.map(cust => ({
      header: `${cust} (箱数)`,
      key: `cust_${cust}`,
      width: Math.max(16, cust.length * 2 + 5)
    }));

    worksheet.columns = [...baseColumns, ...customerColumns];

    // 添加数据并设置样式
    results.forEach((row) => {
      const rowData: Record<string, any> = {
        location: row['库位名称'],
        count: row['箱号数量'],
        available: row['可入库数量'],
      };

      // 填充每列客户对应的箱子数量，为 0 的则填 0，以方便用户做和值计算
      uniqueCustomers.forEach((cust) => {
        const dist = row['客户分布'] || {};
        rowData[`cust_${cust}`] = dist[cust] || 0;
      });

      const excelRow = worksheet.addRow(rowData);

      const count = row['箱号数量'];
      let color = 'FF0F172A'; // 默认：黑色 (Slate 900)
      
      if (count > 18) {
        color = 'FFEF4444'; // 红色 (Red 500)
      } else if (count < 18) {
        color = 'FF10B981'; // 绿色 (Emerald 500)
      }

      // 为整行应用颜色
      excelRow.eachCell((cell) => {
        cell.font = { color: { argb: color }, bold: true };
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
      });
      // 库位名称左对齐
      excelRow.getCell('location').alignment = { vertical: 'middle', horizontal: 'left' };
    });

    // 表头样式
    const headerRow = worksheet.getRow(1);
    headerRow.height = 30;
    headerRow.eachCell((cell) => {
      cell.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 12 };
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FF0F172A' }
      };
      cell.alignment = { vertical: 'middle', horizontal: 'center' };
    });

    // 写入文件
    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    saveAs(blob, `库位统计分析_${new Date().toISOString().split('T')[0]}.xlsx`);
  };

  const exportAlertsToExcel = async () => {
    if (mixedPowerAlerts.length === 0) return;

    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('混档异常明细');

    worksheet.columns = [
      { header: '异常库位', key: 'location', width: 20 },
      { header: '功率档', key: 'power', width: 15 },
      { header: '包含箱号', key: 'boxes', width: 80 }
    ];

    mixedPowerAlerts.forEach((alert) => {
      Object.entries(alert.powerDetails).forEach(([power, boxes]) => {
        const excelRow = worksheet.addRow({
          location: alert.location,
          power: power,
          boxes: (boxes as string[]).join(', ')
        });

        // 样式设置
        excelRow.eachCell((cell) => {
          cell.font = { color: { argb: 'FFEF4444' } }; // 红色字体
          cell.alignment = { vertical: 'middle', horizontal: 'left', wrapText: true };
        });
      });
      // 插入一个空行作为分隔
      worksheet.addRow({});
    });

    // 表头样式
    const headerRow = worksheet.getRow(1);
    headerRow.height = 30;
    headerRow.eachCell((cell) => {
      cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEF4444' } };
      cell.alignment = { vertical: 'middle', horizontal: 'center' };
    });

    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    saveAs(blob, `库存混档异常报告_${new Date().toISOString().split('T')[0]}.xlsx`);
  };

  const exportCustomerAlertsToExcel = async () => {
    if (mixedCustomerAlerts.length === 0) return;

    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('客户混载明细');

    worksheet.columns = [
      { header: '库位名称', key: 'location', width: 20 },
      { header: '客户名称', key: 'customer', width: 30 },
      { header: '包含箱号', key: 'boxes', width: 80 }
    ];

    mixedCustomerAlerts.forEach((alert) => {
      Object.entries(alert.customerDetails).forEach(([customer, boxes]) => {
        const excelRow = worksheet.addRow({
          location: alert.location,
          customer: customer,
          boxes: (boxes as string[]).join(', ')
        });

        excelRow.eachCell((cell) => {
          cell.font = { color: { argb: 'FF2563EB' } }; // 蓝色字体
          cell.alignment = { vertical: 'middle', horizontal: 'left', wrapText: true };
        });
      });
      worksheet.addRow({});
    });

    const headerRow = worksheet.getRow(1);
    headerRow.height = 30;
    headerRow.eachCell((cell) => {
      cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2563EB' } };
      cell.alignment = { vertical: 'middle', horizontal: 'center' };
    });

    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    saveAs(blob, `客户混载预警报告_${new Date().toISOString().split('T')[0]}.xlsx`);
  };

  const exportCustomerStatsToExcel = async () => {
    if (customerStats.length === 0) return;

    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('客户库存数量统计');

    worksheet.columns = [
      { header: '顺序', key: 'index', width: 12 },
      { header: '客户名称', key: 'customerName', width: 45 },
      { header: '占箱数量 (箱)', key: 'boxCount', width: 25 }
    ];

    customerStats.forEach((stat, i) => {
      const excelRow = worksheet.addRow({
        index: i + 1,
        customerName: stat.customerName,
        boxCount: stat.boxCount
      });

      excelRow.eachCell((cell) => {
        cell.font = { color: { argb: 'FF5B21B6' } }; // 紫色
        cell.alignment = { vertical: 'middle', horizontal: 'left' };
      });
      excelRow.getCell('index').alignment = { vertical: 'middle', horizontal: 'center' };
      excelRow.getCell('boxCount').alignment = { vertical: 'middle', horizontal: 'right' };
    });

    const headerRow = worksheet.getRow(1);
    headerRow.height = 30;
    headerRow.eachCell((cell) => {
      cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF7C3AED' } }; // 优雅紫
      cell.alignment = { vertical: 'middle', horizontal: 'center' };
    });

    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    saveAs(blob, `客户库存清单报告_${new Date().toISOString().split('T')[0]}.xlsx`);
  };

  const generateLabelsPDF = (locationName: string) => {
    const doc = new jsPDF({
      orientation: 'portrait',
      unit: 'mm',
      format: 'letter'
    });

    const pageWidth = 215.9;
    const pageHeight = 279.4;
    const labelHeight = pageHeight / 3;

    for (let p = 0; p < 3; p++) {
      if (p > 0) doc.addPage();
      
      for (let l = 0; l < 3; l++) {
        const labelIndex = p * 3 + l + 1;
        const suffix = `-${labelIndex}`;
        const yOffset = l * labelHeight;

        // Configuration
        const baseFontSize = 140; // Extremely large for the main location
        const suffixFontSize = 60; // Smaller for the index
        const dateFontSize = 14;

        // Set font for calculations
        doc.setFont('helvetica', 'bold');
        
        // Calculate widths to center the combined string
        doc.setFontSize(baseFontSize);
        const baseWidth = doc.getTextWidth(locationName);
        doc.setFontSize(suffixFontSize);
        const suffixWidth = doc.getTextWidth(suffix);
        
        const totalWidth = baseWidth + suffixWidth + 2; // +2 for a tiny spacing
        const startX = (pageWidth - totalWidth) / 2;
        const centerY = yOffset + (labelHeight / 2) + 10;

        // Draw Location Name (Large)
        doc.setTextColor(15, 23, 42); // Slate 900
        doc.setFontSize(baseFontSize);
        doc.text(locationName, startX, centerY);
        
        // Draw Suffix (Smaller)
        doc.setFontSize(suffixFontSize);
        doc.text(suffix, startX + baseWidth + 2, centerY);
        
        // Date text - Slightly adjusted secondary text
        doc.setFontSize(dateFontSize);
        doc.setFont('helvetica', 'normal');
        doc.setTextColor(100);
        const dateText = `Printed: ${new Date().toLocaleDateString()} ${new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })}`;
        const dateWidth = doc.getTextWidth(dateText);
        doc.text(dateText, (pageWidth - dateWidth) / 2, yOffset + labelHeight - 15);

        // Separation line
        doc.setDrawColor(220);
        doc.setLineWidth(0.3);
        if (l < 2) {
          doc.line(20, yOffset + labelHeight, pageWidth - 20, yOffset + labelHeight);
        }
      }
    }

    doc.save(`标签_${locationName}_L9.pdf`);
  };

  const reset = () => {
    setData([]);
    setResults([]);
    setMixedPowerAlerts([]);
    setMixedCustomerAlerts([]);
    setOrderCustomerMap({});
    setMappingFileName('');
    setCustomerStats([]);
    setFileName('');
    setError(null);
  };

  return (
    <div className="flex h-screen w-full bg-[#F8FAFC] text-slate-800 font-sans overflow-hidden">
      {/* Left Navigation Rail */}
      <aside className="w-64 bg-slate-900 flex flex-col shrink-0">
        <div className="p-6 flex items-center gap-3">
          <div className="w-8 h-8 bg-blue-500 rounded-lg flex items-center justify-center shadow-lg shadow-blue-500/20">
            <BarChart3 className="w-5 h-5 text-white" />
          </div>
          <span className="text-white font-bold tracking-tight uppercase">Inv Analyst</span>
        </div>
        
        <nav className="mt-4 flex-1">
          <div className="px-6 py-2 text-[10px] font-bold text-slate-500 uppercase tracking-widest mb-2 opacity-50">数据管理</div>
          <button 
            onClick={reset}
            className={cn(
              "w-full flex items-center gap-3 px-6 py-3 transition-all text-sm font-medium relative group",
              !data.length ? "bg-blue-600 text-white" : "text-slate-400 hover:text-white hover:bg-slate-800"
            )}
          >
            {!data.length && <span className="absolute left-0 top-0 bottom-0 w-1 bg-white animate-pulse" />}
            <Upload className="w-4 h-4" />
            重置分析沙盒
          </button>
        </nav>

        <div className="p-6 border-t border-slate-800 text-[10px] text-slate-500 uppercase font-bold tracking-widest">
          系统状态: <span className="text-emerald-500">运行中</span>
        </div>
      </aside>

      {/* Main Content Area */}
      <main className="flex-1 flex flex-col overflow-hidden">
        {/* Top Header */}
        <header className="h-16 bg-white border-b border-slate-200 flex items-center justify-between px-8 shrink-0">
          <div className="flex items-center gap-2">
            <h2 className="text-lg font-bold text-slate-800">库存数据统计分析</h2>
            <span className="px-2 py-0.5 bg-slate-100 text-[10px] text-slate-400 rounded-full font-bold uppercase tracking-wider">v1.2.1</span>
          </div>
          
          <div className="flex items-center gap-6">
            {fileName && (
              <div className="text-right hidden sm:block">
                <p className="text-[10px] uppercase font-bold text-slate-400 leading-none mb-1">当前操作文件</p>
                <p className="text-sm font-bold text-slate-700">{fileName}</p>
              </div>
            )}
            <div className="h-10 w-px bg-slate-100 hidden sm:block" />
            <div className="w-10 h-10 rounded-full bg-slate-50 border border-slate-200 flex items-center justify-center text-slate-400 hover:text-slate-600 transition-colors cursor-help">
              <FileText className="w-5 h-5" />
            </div>
          </div>
        </header>

        <div className="p-8 flex-1 overflow-auto bg-[#F8FAFC]">
          <div className="max-w-7xl mx-auto grid grid-cols-12 gap-8 h-full">
            
            {/* Left Control Column */}
            <div className="col-span-12 xl:col-span-4 flex flex-col gap-6">
              
              {/* Card 1: Main Inventory Data Upload */}
              <div className="bg-white p-6 rounded-3xl border border-slate-200 shadow-sm flex flex-col gap-4">
                <div className="flex items-center justify-between pb-3 border-b border-slate-100">
                  <h4 className="text-xs font-bold uppercase tracking-wider text-slate-500 flex items-center gap-2">
                    <span className="w-5 h-5 rounded-full bg-blue-100 text-blue-600 flex items-center justify-center font-bold text-[10px]">1</span>
                    主库存数据源
                  </h4>
                  {data.length > 0 && (
                    <span className="px-2 py-0.5 bg-emerald-50 text-emerald-600 rounded-full text-[9px] font-bold border border-emerald-100 animate-pulse">
                      正常载入
                    </span>
                  )}
                </div>

                {!data.length ? (
                  <motion.div 
                    initial={{ opacity: 0, y: 10 }}
                    animate={{ opacity: 1, y: 0 }}
                    onDragOver={(e) => e.preventDefault()}
                    onDrop={handleDrop}
                    className="p-8 rounded-2xl border-2 border-dashed border-slate-200 flex flex-col items-center justify-center text-center relative group hover:border-blue-500 hover:bg-blue-50/10 transition-all duration-300 min-h-[160px]"
                  >
                    <input 
                      type="file" 
                      accept=".xlsx,.xls" 
                      onChange={handleFileUpload}
                      className="absolute inset-0 opacity-0 cursor-pointer z-10"
                    />
                    <div className="w-12 h-12 bg-blue-50 rounded-2xl flex items-center justify-center mb-4 text-blue-600 group-hover:scale-110 transition-transform">
                      <Upload className="w-6 h-6" />
                    </div>
                    <h5 className="font-bold text-sm text-slate-800">拖拽主库存 Excel 到此处</h5>
                    <p className="text-[10px] text-slate-400 mt-1">或点击选择本地文件格式 (.xlsx, .xls)</p>
                    
                    {isProcessing && (
                      <div className="absolute inset-0 bg-white/95 backdrop-blur-sm rounded-2xl flex flex-col items-center justify-center z-20">
                        <Loader2 className="w-8 h-8 text-blue-600 animate-spin mb-2" />
                        <p className="text-[10px] font-bold text-slate-800 tracking-wider">正在解析数据...</p>
                      </div>
                    )}
                  </motion.div>
                ) : (
                  <div className="bg-slate-900 rounded-2xl p-5 text-white shadow-md relative overflow-hidden">
                    <div className="absolute -right-6 -bottom-6 w-20 h-20 bg-white/5 rounded-full" />
                    <div className="space-y-4">
                      <div>
                        <p className="text-[9px] text-slate-500 uppercase tracking-widest font-bold">已连接库存文件</p>
                        <p className="text-sm font-bold truncate mt-1 text-slate-100">{fileName}</p>
                      </div>
                      <div className="grid grid-cols-2 gap-2 border-t border-slate-800 pt-3">
                        <div>
                          <p className="text-[9px] text-slate-500 uppercase">数据吞吐行数</p>
                          <p className="text-lg font-bold font-mono tracking-tight text-blue-400">{data.length}</p>
                        </div>
                        <div>
                          <p className="text-[9px] text-slate-500 uppercase">已锁定库位数</p>
                          <p className="text-lg font-bold font-mono tracking-tight text-emerald-400">{results.length}</p>
                        </div>
                      </div>
                      <div className="flex gap-2">
                        <button 
                          onClick={exportToExcel}
                          className="flex-1 flex items-center justify-center gap-1.5 bg-emerald-600 hover:bg-emerald-500 text-white font-bold py-2 rounded-xl text-xs transition-colors shadow-md active:scale-95"
                        >
                          <Download className="w-3.5 h-3.5" />
                          导出库位表
                        </button>
                        <button 
                          onClick={() => { setData([]); setResults([]); setFileName(''); }}
                          className="px-3 bg-slate-800 hover:bg-slate-700 text-slate-400 hover:text-white rounded-xl text-xs transition-colors"
                          title="置空当前库存数据"
                        >
                          重置
                        </button>
                      </div>
                    </div>
                  </div>
                )}
              </div>

              {/* Card 2: Work Order -> Customer Relations Mapping */}
              <div className="bg-white p-6 rounded-3xl border border-slate-200 shadow-sm flex flex-col gap-4">
                <div className="flex items-center justify-between pb-3 border-b border-slate-100">
                  <h4 className="text-xs font-bold uppercase tracking-wider text-slate-500 flex items-center gap-2">
                    <span className="w-5 h-5 rounded-full bg-violet-100 text-violet-600 flex items-center justify-center font-bold text-[10px]">2</span>
                    工单与客户对照表
                  </h4>
                  {mappingFileName && (
                    <span className="px-2 py-0.5 bg-violet-50 text-violet-600 rounded-full text-[9px] font-bold border border-violet-100 animate-pulse">
                      对照正常
                    </span>
                  )}
                </div>

                {!mappingFileName ? (
                  <motion.div 
                    initial={{ opacity: 0, y: 10 }}
                    animate={{ opacity: 1, y: 0 }}
                    onDragOver={(e) => e.preventDefault()}
                    onDrop={handleMappingDrop}
                    className="p-8 rounded-2xl border-2 border-dashed border-slate-200 flex flex-col items-center justify-center text-center relative group hover:border-violet-500 hover:bg-violet-50/10 transition-all duration-300 min-h-[160px]"
                  >
                    <input 
                      type="file" 
                      accept=".xlsx,.xls" 
                      onChange={handleMappingFileUpload}
                      className="absolute inset-0 opacity-0 cursor-pointer z-10"
                    />
                    <div className="w-12 h-12 bg-violet-50 rounded-2xl flex items-center justify-center mb-4 text-violet-600 group-hover:scale-110 transition-transform">
                      <Package className="w-6 h-6 animate-pulse" />
                    </div>
                    <h5 className="font-bold text-sm text-slate-800">拖拽工单-客户对照表此处</h5>
                    <p className="text-[10px] text-slate-400 mt-1">建立库存销售单号/工单号与客户名称对应关系</p>
                  </motion.div>
                ) : (
                  <div className="bg-violet-950 rounded-2xl p-5 text-white shadow-md relative overflow-hidden">
                    <div className="absolute -right-6 -bottom-6 w-20 h-20 bg-white/5 rounded-full" />
                    <div className="space-y-4">
                      <div>
                        <p className="text-[9px] text-violet-400 uppercase tracking-widest font-bold">已连接对照字典</p>
                        <p className="text-sm font-bold truncate mt-1 text-slate-100">{mappingFileName}</p>
                      </div>
                      <div className="grid grid-cols-2 gap-2 border-t border-violet-900 pt-3">
                        <div>
                          <p className="text-[9px] text-violet-400 uppercase">已识别关联规则</p>
                          <p className="text-lg font-bold font-mono tracking-tight text-violet-300">{Object.keys(orderCustomerMap).length} 条</p>
                        </div>
                        <div>
                          <p className="text-[9px] text-violet-400 uppercase">已翻译关联客户</p>
                          <p className="text-lg font-bold font-mono tracking-tight text-violet-300">{customerStats.filter(c => c.customerName !== '未映射客户').length} 位</p>
                        </div>
                      </div>
                      <div className="flex gap-2">
                        {customerStats.length > 0 && (
                          <button 
                            onClick={exportCustomerStatsToExcel}
                            className="flex-1 flex items-center justify-center gap-1.5 bg-violet-750 hover:bg-violet-650 text-white font-bold py-2 rounded-xl text-xs transition-colors border border-violet-600 shadow-md active:scale-95"
                          >
                            <Download className="w-3.5 h-3.5" />
                            导出客户统计
                          </button>
                        )}
                        <button 
                          onClick={clearMapping}
                          className="px-3 bg-violet-900 hover:bg-violet-800 text-violet-300 hover:text-white rounded-xl text-xs transition-colors"
                          title="置空对应关系对照字典"
                        >
                          清除
                        </button>
                      </div>
                    </div>
                  </div>
                )}
              </div>

              {/* Requirements Guidelines */}
              <div className="bg-white p-7 rounded-3xl border border-slate-200 shadow-sm relative overflow-hidden group">
                <div className="absolute top-0 right-0 w-24 h-24 bg-blue-50 rounded-full -mr-12 -mt-12 group-hover:scale-110 transition-transform duration-700 opacity-60" />
                <h3 className="text-xs font-bold uppercase tracking-widest text-slate-400 mb-6 flex items-center gap-2">
                  <MapPin className="w-3 h-3" />
                  Excel 格式要求指引
                </h3>
                <div className="space-y-4">
                  <div className="p-3 bg-slate-50 rounded-2xl border border-slate-100">
                    <span className="text-[10px] font-bold text-blue-600 uppercase block mb-1">库存源数据列名需包含：</span>
                    <ul className="text-[10px] text-slate-500 list-disc list-inside space-y-0.5">
                      <li><b>库位名称</b> (如 C36-1)</li>
                      <li><b>箱号</b> (唯一标识符)</li>
                      <li><b>功率档</b> / <b>工单号</b> (可选)</li>
                    </ul>
                  </div>
                  <div className="p-3 bg-slate-50 rounded-2xl border border-slate-100">
                    <span className="text-[10px] font-bold text-violet-600 block mb-1">对照关系表列名需包含：</span>
                    <ul className="text-[10px] text-slate-500 list-disc list-inside space-y-0.5">
                      <li>含有 <b>工单/销售单</b> 关键字的列</li>
                      <li>含有 <b>客户/Client</b> 关键字的列</li>
                    </ul>
                  </div>
                </div>
              </div>
            </div>

            {/* Right Results Column */}
            <div className="col-span-12 xl:col-span-8 flex flex-col gap-6 h-full min-h-[600px]">
              {/* Summary Insights */}
              {results.length > 0 && (
                <div className="grid grid-cols-1 lg:grid-cols-3 gap-6 shrink-0">
                  {/* Warehouse Overview Info */}
                  <motion.div 
                    initial={{ opacity: 0, y: 10 }}
                    animate={{ opacity: 1, y: 0 }}
                    transition={{ delay: 0.1 }}
                    className="bg-white rounded-3xl p-6 border border-slate-200 shadow-sm"
                  >
                    <div className="flex flex-col gap-4">
                      <div className="flex items-center justify-between pb-4 border-b border-slate-50">
                        <h4 className="text-sm font-bold text-slate-800 flex items-center gap-2">
                          <BarChart3 className="w-4 h-4 text-blue-500" />
                          库存概览
                        </h4>
                        <div className="flex gap-2">
                          <span className="px-2 py-0.5 bg-blue-50 text-blue-600 text-[10px] font-bold rounded-full">
                            库位: {results.length}
                          </span>
                        </div>
                      </div>
                      
                      <div className="grid grid-cols-2 gap-4">
                        <div className="bg-slate-50 rounded-2xl p-3">
                          <div className="text-[10px] text-slate-400 font-bold mb-1">当前箱号总数</div>
                          <div className="text-lg font-black text-slate-900">
                            {results.reduce((acc, curr) => acc + curr['箱号数量'], 0)}
                          </div>
                        </div>
                        <div className="bg-emerald-50 rounded-2xl p-3">
                          <div className="text-[10px] text-emerald-600 font-bold mb-1">剩余可入位</div>
                          <div className="text-lg font-black text-emerald-700">
                            {results.reduce((acc, curr) => acc + (curr['可入库数量'] > 0 ? curr['可入库数量'] : 0), 0)}
                          </div>
                        </div>
                      </div>

                      <div className="space-y-2">
                        <div className="flex items-center justify-between text-[10px] font-bold">
                          <span className="text-slate-400 uppercase tracking-wider flex items-center gap-1">
                            <CheckCircle2 className="w-3 h-3 text-emerald-500" />
                            完整可用库位 ({results.filter(r => r['箱号数量'] === 0).length})
                          </span>
                        </div>
                        <div className="max-h-16 overflow-y-auto pr-2 scrollbar-thin flex flex-wrap gap-1.5 text-[9px]">
                          {results.filter(r => r['箱号数量'] === 0).length > 0 ? (
                            results.filter(r => r['箱号数量'] === 0).map(r => (
                              <button 
                                key={r['库位名称']} 
                                onClick={() => generateLabelsPDF(r['库位名称'])}
                                className="px-2 py-0.5 bg-white text-slate-500 rounded-md border border-slate-200 font-mono shadow-sm hover:border-emerald-400 hover:text-emerald-600 hover:shadow-md transition-all active:scale-95 group/tag inline-flex items-center gap-1"
                                title="点击生成 PDF 标签"
                              >
                                {r['库位名称']}
                                <FileText className="w-2.5 h-2.5 opacity-0 group-hover/tag:opacity-100 transition-opacity" />
                              </button>
                            ))
                          ) : (
                            <span className="text-slate-400 italic">暂无空置库位</span>
                          )}
                        </div>
                      </div>
                    </div>
                  </motion.div>

                  {/* Mixed Power Alerts Summary */}
                  <motion.div 
                    initial={{ opacity: 0, y: 10 }}
                    animate={{ opacity: 1, y: 0 }}
                    transition={{ delay: 0.2 }}
                    className={cn(
                      "rounded-3xl p-6 border shadow-sm",
                      mixedPowerAlerts.length > 0 ? "bg-red-50 border-red-100" : "bg-white border-slate-200"
                    )}
                  >
                    <div className="flex items-center justify-between mb-4">
                      <div className="flex items-center gap-2">
                        <h4 className="text-sm font-bold flex items-center gap-2 text-nowrap">
                          <AlertCircle className={cn("w-4 h-4", mixedPowerAlerts.length > 0 ? "text-red-500" : "text-slate-400")} />
                          功率混档监测
                        </h4>
                        {mixedPowerAlerts.length > 0 && (
                          <button 
                            onClick={exportAlertsToExcel}
                            className="p-1.5 hover:bg-red-100 text-red-600 rounded-lg transition-colors group/btn"
                            title="导出混档异常"
                          >
                            <Download className="w-3.5 h-3.5" />
                          </button>
                        )}
                      </div>
                      <span className={cn("text-xl font-black", mixedPowerAlerts.length > 0 ? "text-red-600" : "text-slate-400")}>
                        {mixedPowerAlerts.length}
                      </span>
                    </div>
                    <div className="max-h-24 overflow-y-auto pr-2 scrollbar-thin">
                      {mixedPowerAlerts.length > 0 ? (
                        <div className="space-y-3">
                          {mixedPowerAlerts.map((alert, i) => (
                            <div key={i} className="flex flex-col p-3 bg-white rounded-xl border border-red-200 text-[10px] shadow-sm">
                              <div className="font-bold text-slate-800 mb-2 flex items-center gap-1">
                                <Box className="w-3 h-3 text-red-400" />
                                {alert.location}
                              </div>
                              <div className="space-y-1">
                                {Object.keys(alert.powerDetails).map(p => (
                                  <div key={p} className="flex items-center justify-between text-slate-500">
                                    <span>功率 {p}</span>
                                    <span className="font-mono text-red-600 bg-red-50 px-1.5 rounded">{alert.powerDetails[p].length}箱</span>
                                  </div>
                                ))}
                              </div>
                            </div>
                          ))}
                        </div>
                      ) : (
                        <p className="text-slate-400 text-xs italic">库位功率一致性良好</p>
                      )}
                    </div>
                  </motion.div>

                  {/* Mixed Customer Alerts Summary */}
                  <motion.div 
                    initial={{ opacity: 0, y: 10 }}
                    animate={{ opacity: 1, y: 0 }}
                    transition={{ delay: 0.3 }}
                    className={cn(
                      "rounded-3xl p-6 border shadow-sm",
                      mixedCustomerAlerts.length > 0 ? "bg-blue-50 border-blue-100" : "bg-white border-slate-200"
                    )}
                  >
                    <div className="flex items-center justify-between mb-4">
                      <div className="flex items-center gap-2">
                        <h4 className="text-sm font-bold flex items-center gap-2 text-nowrap">
                          <History className={cn("w-4 h-4", mixedCustomerAlerts.length > 0 ? "text-blue-500" : "text-slate-400")} />
                          客户混载监测
                        </h4>
                        {mixedCustomerAlerts.length > 0 && (
                          <button 
                            onClick={exportCustomerAlertsToExcel}
                            className="p-1.5 hover:bg-blue-100 text-blue-600 rounded-lg transition-colors group/btn"
                            title="导出客户混载"
                          >
                            <Download className="w-3.5 h-3.5" />
                          </button>
                        )}
                      </div>
                      <span className={cn("text-xl font-black", mixedCustomerAlerts.length > 0 ? "text-blue-600" : "text-slate-400")}>
                        {mixedCustomerAlerts.length}
                      </span>
                    </div>
                    <div className="max-h-24 overflow-y-auto pr-2 scrollbar-thin">
                      {mixedCustomerAlerts.length > 0 ? (
                        <div className="space-y-3">
                          {mixedCustomerAlerts.map((alert, i) => (
                            <div key={i} className="flex flex-col p-3 bg-white rounded-xl border border-blue-200 text-[10px] shadow-sm">
                              <div className="font-bold text-slate-800 mb-2 flex items-center gap-1">
                                <Package className="w-3 h-3 text-blue-400" />
                                {alert.location}
                              </div>
                              <div className="space-y-1">
                                {Object.keys(alert.customerDetails).map(c => (
                                  <div key={c} className="flex items-center justify-between text-slate-500">
                                    <span className="truncate max-w-[100px]" title={c}>{c}</span>
                                    <span className="font-mono text-blue-600 bg-blue-50 px-1.5 rounded">{alert.customerDetails[c].length}箱</span>
                                  </div>
                                ))}
                              </div>
                            </div>
                          ))}
                        </div>
                      ) : (
                        <p className="text-slate-400 text-xs italic">库位客户单一，未发现混载</p>
                      )}
                    </div>
                  </motion.div>
                </div>
              )}

              {/* Main Detail Table */}
              <div className="bg-white rounded-3xl border border-slate-200 shadow-sm overflow-hidden flex flex-col flex-1">
                <div className="p-6 border-b border-slate-100 flex flex-col gap-4 sm:flex-row sm:items-center sm:justify-between bg-white z-20">
                  <div>
                    <h3 className="font-bold text-xl text-slate-900 tracking-tight">分析明细看板</h3>
                    <p className="text-xs text-slate-400 font-medium">
                      {activeTab === 'location' ? '统计并展示各库位上的箱号负载与余位数量' : '对照翻译显示各客户包含的总箱数及库位分布'}
                    </p>
                  </div>

                  <div className="flex items-center gap-4 self-start sm:self-auto">
                    {/* Tab Segment Controller */}
                    <div className="flex bg-slate-100 p-1 rounded-xl">
                      <button 
                        onClick={() => setActiveTab('location')}
                        className={cn(
                          "px-4 py-1.5 text-xs font-bold rounded-lg transition-all flex items-center gap-1.5",
                          activeTab === 'location' ? "bg-white text-slate-950 shadow-sm" : "text-slate-550 hover:text-slate-900"
                        )}
                      >
                        <MapPin className="w-3.5 h-3.5" />
                        按库位
                      </button>
                      <button 
                        onClick={() => setActiveTab('customer')}
                        className={cn(
                          "px-4 py-1.5 text-xs font-bold rounded-lg transition-all flex items-center gap-1.5",
                          activeTab === 'customer' ? "bg-white text-slate-950 shadow-sm" : "text-slate-550 hover:text-slate-900"
                        )}
                      >
                        <Package className="w-3.5 h-3.5" />
                        按客户
                      </button>
                    </div>

                    {(activeTab === 'location' ? results.length > 0 : customerStats.length > 0) && (
                      <div className="flex items-center gap-3">
                        <div className="w-px h-8 bg-slate-100 hidden sm:block" />
                        <button 
                          onClick={activeTab === 'location' ? exportToExcel : exportCustomerStatsToExcel}
                          className="p-2.5 bg-slate-50 hover:bg-slate-100 text-slate-605 rounded-xl border border-slate-200 transition-colors flex items-center gap-1.5 shrink-0 text-xs font-bold"
                          title="导出当前分表为 Excel"
                        >
                          <Download className="w-4 h-4" />
                          <span>导出 Excel</span>
                        </button>
                      </div>
                    )}
                  </div>
                </div>

                <div className="flex-1 overflow-auto">
                  <table className="w-full border-collapse">
                    {activeTab === 'location' ? (
                      <>
                        <thead className="sticky top-0 bg-white/95 backdrop-blur-sm z-10 border-b border-slate-100 shadow-sm shadow-slate-100/50">
                          <tr className="text-left">
                            <th className="pl-8 pr-4 py-5 text-[10px] font-bold text-slate-400 uppercase tracking-[0.2em] w-16">顺序</th>
                            <th className="px-6 py-5 text-[10px] font-bold text-slate-400 uppercase tracking-[0.2em]">库位名称</th>
                            <th className="px-6 py-5 text-[10px] font-bold text-slate-400 uppercase tracking-[0.2em] text-center">箱号数量</th>
                            <th className="px-6 py-5 text-[10px] font-bold text-slate-400 uppercase tracking-[0.2em] text-right pr-12">可入库数量 (标准18)</th>
                          </tr>
                        </thead>
                        <tbody className="divide-y divide-slate-50">
                          {results.length > 0 ? (
                            results.map((row, i) => (
                              <motion.tr 
                                key={i}
                                initial={{ opacity: 0, y: 10 }}
                                animate={{ opacity: 1, y: 0 }}
                                transition={{ delay: i * 0.015, duration: 0.3 }}
                                className="hover:bg-slate-50/50 transition-all duration-200 group"
                              >
                                <td className="pl-8 pr-4 py-4 text-[10px] font-mono font-bold text-slate-300 group-hover:text-blue-400 transition-colors">
                                  {String(i + 1).padStart(2, '0')}
                                </td>
                                <td className="px-6 py-4">
                                  <div className="flex items-center gap-3">
                                    <div className={cn(
                                      "w-2 h-2 rounded-full scale-50 group-hover:scale-100 transition-all shrink-0",
                                      row['箱号数量'] > 18 ? "bg-red-500 shadow-[0_0_8px_rgba(239,68,68,0.5)]" : 
                                      row['箱号数量'] < 18 ? "bg-emerald-500 shadow-[0_0_8px_rgba(16,185,129,0.5)]" : 
                                      "bg-slate-900 shadow-[0_0_8px_rgba(15,23,42,0.3)]"
                                    )} />
                                    <div className="flex flex-col">
                                      <span className={cn(
                                        "text-sm font-bold transition-colors",
                                        row['箱号数量'] > 18 ? "text-red-600" : 
                                        row['箱号数量'] < 18 ? "text-emerald-700" :
                                        "text-slate-900"
                                      )}>
                                        {row['库位名称']}
                                      </span>
                                      {row['客户分布'] && Object.keys(row['客户分布']).length > 0 && (
                                        <span className="text-[10px] text-slate-400 mt-0.5" title={
                                          Object.entries(row['客户分布'])
                                            .map(([cust, count]) => `${cust} (${count}箱)`)
                                            .join(', ')
                                        }>
                                          客户分布: {
                                            Object.entries(row['客户分布'])
                                              .map(([cust, count]) => `${cust} (${count}箱)`)
                                              .join(', ')
                                          }
                                        </span>
                                      )}
                                    </div>
                                  </div>
                                </td>
                                <td className="px-6 py-4 text-center">
                                  <span className={cn(
                                    "text-sm font-mono font-bold tabular-nums px-3 py-1 rounded-full",
                                    row['箱号数量'] > 18 ? "bg-red-50 text-red-600" : 
                                    row['箱号数量'] < 18 ? "bg-emerald-50 text-emerald-600" : 
                                    "bg-slate-100 text-slate-900"
                                  )}>
                                    {row['箱号数量']}
                                  </span>
                                </td>
                                <td className="px-6 py-4 text-right pr-12">
                                  <span className={cn(
                                    "text-sm font-mono font-bold tabular-nums transition-colors",
                                    row['可入库数量'] < 0 ? "text-red-500" : 
                                    row['可入库数量'] > 0 ? "text-emerald-500" : 
                                    "text-slate-900"
                                  )}>
                                    {row['可入库数量'] > 0 ? `+${row['可入库数量']}` : row['可入库数量']}
                                  </span>
                                </td>
                              </motion.tr>
                            ))
                          ) : (
                            <tr>
                              <td colSpan={4} className="px-8 py-32 text-center">
                                <div className="flex flex-col items-center justify-center">
                                  <div className="w-16 h-16 bg-slate-50 rounded-full flex items-center justify-center mb-4 text-slate-200">
                                    <TableIcon className="w-8 h-8" />
                                  </div>
                                  <p className="text-sm font-bold text-slate-400 tracking-wide">等待分析文件载入...</p>
                                  <p className="text-[10px] text-slate-300 uppercase font-medium mt-1">请上传包含箱号信息的库存 Excel</p>
                                </div>
                              </td>
                            </tr>
                          )}
                        </tbody>
                      </>
                    ) : (
                      <>
                        <thead className="sticky top-0 bg-white/95 backdrop-blur-sm z-10 border-b border-slate-100 shadow-sm shadow-slate-100/50">
                          <tr className="text-left">
                            <th className="pl-8 pr-4 py-5 text-[10px] font-bold text-slate-400 uppercase tracking-[0.2em] w-16">顺序</th>
                            <th className="px-6 py-5 text-[10px] font-bold text-slate-400 uppercase tracking-[0.2em]">客户名称</th>
                            <th className="px-6 py-5 text-[10px] font-bold text-slate-400 uppercase tracking-[0.2em] text-center">占箱数</th>
                            <th className="px-6 py-5 text-[10px] font-bold text-slate-400 uppercase tracking-[0.2em] pr-12">占用库位明细</th>
                          </tr>
                        </thead>
                        <tbody className="divide-y divide-slate-50">
                          {customerStats.length > 0 ? (
                            customerStats.map((row, i) => (
                              <motion.tr 
                                key={i}
                                initial={{ opacity: 0, y: 10 }}
                                animate={{ opacity: 1, y: 0 }}
                                transition={{ delay: i * 0.015, duration: 0.3 }}
                                className="hover:bg-slate-50/50 transition-all duration-200 group"
                              >
                                <td className="pl-8 pr-4 py-4 text-[10px] font-mono font-bold text-slate-300 group-hover:text-violet-400 transition-colors">
                                  {String(i + 1).padStart(2, '0')}
                                </td>
                                <td className="px-6 py-4">
                                  <span className="text-sm font-bold text-slate-900 block">
                                    {row.customerName}
                                  </span>
                                </td>
                                <td className="px-6 py-4 text-center">
                                  <span className="text-sm font-mono font-bold bg-violet-50 text-violet-600 px-3 py-1 rounded-full border border-violet-100">
                                    {row.boxCount} 箱
                                  </span>
                                </td>
                                <td className="px-6 py-4 pr-12">
                                  <div className="flex flex-wrap gap-1.5 max-h-24 overflow-y-auto pr-1">
                                    {row.locations.map(loc => (
                                      <span key={loc} className="px-2 py-0.5 bg-slate-100 text-slate-600 text-[10px] font-mono rounded font-bold border border-slate-200">
                                        {loc}
                                      </span>
                                    ))}
                                  </div>
                                </td>
                              </motion.tr>
                            ))
                          ) : (
                            <tr>
                              <td colSpan={4} className="px-8 py-32 text-center">
                                <div className="flex flex-col items-center justify-center max-w-sm mx-auto">
                                  <div className="w-16 h-16 bg-violet-55 text-violet-500 rounded-full flex items-center justify-center mb-4">
                                    <Package className="w-8 h-8" />
                                  </div>
                                  <p className="text-sm font-bold text-slate-600">未发现客户关联统计数据</p>
                                  <p className="text-xs text-slate-400 mt-2 leading-relaxed">
                                    欲按“客户”维度归类分析箱数，请在左侧上传带有工单及客户信息的 <b>对照表</b>。系统支持在拖拽上传后实时计算并加载该明细视图。
                                  </p>
                                </div>
                              </td>
                            </tr>
                          )}
                        </tbody>
                      </>
                    )}
                  </table>
                </div>

                <div className="p-6 bg-slate-50/50 border-t border-slate-100 flex items-center justify-between">
                  <div className="flex items-center gap-1.5 opacity-60">
                    <CheckCircle2 className="w-3.5 h-3.5 text-emerald-500" />
                    <span className="text-[10px] font-bold text-slate-500 uppercase tracking-widest">分析引擎 v4.0</span>
                  </div>
                  <p className="text-[10px] text-slate-400 font-medium">
                    生成于 {new Date().toLocaleTimeString()} • 本地沙盒环境处理
                  </p>
                </div>
              </div>
            </div>
          </div>
        </div>
      </main>

      {/* Error Boundary Toast */}
      <AnimatePresence>
        {error && (
          <motion.div 
            initial={{ opacity: 0, y: 30 }}
            animate={{ opacity: 1, y: 0 }}
            exit={{ opacity: 0, scale: 0.95 }}
            className="fixed bottom-10 right-10 bg-slate-900 text-white px-8 py-5 rounded-2xl shadow-[0_20px_50px_rgba(0,0,0,0.3)] flex items-center gap-4 z-50 border border-white/10 backdrop-blur-xl"
          >
            <div className="w-10 h-10 bg-red-500 rounded-full flex items-center justify-center shrink-0 shadow-lg shadow-red-500/20">
              <AlertCircle className="w-5 h-5 text-white" />
            </div>
            <div>
              <p className="text-[10px] font-bold text-red-400 uppercase tracking-widest mb-0.5">分析错误</p>
              <p className="text-sm font-bold">{error}</p>
            </div>
            <button onClick={() => setError(null)} className="ml-4 p-2 hover:bg-white/10 rounded-lg transition-colors text-slate-400">
              <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M6 18L18 6M6 6l12 12" /></svg>
            </button>
          </motion.div>
        )}
      </AnimatePresence>
    </div>
  );
}

