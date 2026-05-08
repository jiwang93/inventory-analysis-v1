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
}

interface MixedPowerAlert {
  location: string;
  powerDetails: Record<string, string[]>; 
}

interface MixedOrderAlert {
  location: string;
  orderDetails: Record<string, string[]>; // OrderNo -> List of BoxNumbers
}

export default function App() {
  const STANDARD_CAPACITY = 18;
  const [data, setData] = useState<InventoryRow[]>([]);
  const [results, setResults] = useState<AnalysisResult[]>([]);
  const [mixedPowerAlerts, setMixedPowerAlerts] = useState<MixedPowerAlert[]>([]);
  const [mixedOrderAlerts, setMixedOrderAlerts] = useState<MixedOrderAlert[]>([]);
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
        const rawData = XLSX.utils.sheet_to_json(worksheet) as any[];

        if (!rawData || rawData.length === 0) {
          throw new Error('表格中没有数据');
        }

        // Validate columns
        const requiredColumns = ['箱号', '库位名称'];
        const firstRow = rawData[0];
        const missing = requiredColumns.filter(col => !(col in firstRow));

        if (missing.length > 0) {
          throw new Error(`缺少必要列: ${missing.join(', ')}`);
        }

        setData(rawData);
        performAnalysis(rawData);
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

  const performAnalysis = (rows: InventoryRow[]) => {
    const locationMap: Record<string, Set<string | number>> = {};
    // 功率检测：库位 -> { 功率档 -> [箱号] }
    const locationPowerData: Record<string, Record<string, Set<string>>> = {};
    // 工单监测：库位 -> { 工单号 -> [箱号] }
    const locationOrderData: Record<string, Record<string, Set<string>>> = {};

    // 1. 初始化固定的 C-H (1-60) 库位
    const zones = ['C', 'D', 'E', 'F', 'G', 'H'];
    zones.forEach(zone => {
      for (let i = 1; i <= 60; i++) {
        const locName = `${zone}${String(i).padStart(2, '0')}`;
        locationMap[locName] = new Set();
        locationPowerData[locName] = {};
        locationOrderData[locName] = {};
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
      
      if (!locationMap[loc]) {
        locationMap[loc] = new Set();
        locationPowerData[loc] = {};
        locationOrderData[loc] = {};
      }

      if (box !== undefined && box !== null && String(box).trim() !== '') {
        const boxStr = String(box);
        locationMap[loc].add(boxStr);

        // 记录库位下的功率分布
        if (!locationPowerData[loc][power]) {
          locationPowerData[loc][power] = new Set();
        }
        locationPowerData[loc][power].add(boxStr);

        // 记录库位下的工单分布
        if (!locationOrderData[loc][orderNo]) {
          locationOrderData[loc][orderNo] = new Set();
        }
        locationOrderData[loc][orderNo].add(boxStr);
      }
    });

    // 汇总分析结果
    const analysisResults: AnalysisResult[] = Object.entries(locationMap).map(([location, boxSet]) => {
      const boxCount = boxSet.size;
      return {
        '库位名称': location,
        '箱号数量': boxCount,
        '可入库数量': STANDARD_CAPACITY - boxCount
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

    // 处理工单混载警告
    const orderAlerts: MixedOrderAlert[] = [];
    Object.entries(locationOrderData).forEach(([loc, orders]) => {
      const orderNos = Object.keys(orders);
      if (orderNos.length > 1) {
        const orderDetails: Record<string, string[]> = {};
        orderNos.forEach(o => {
          orderDetails[o] = Array.from(orders[o]);
        });
        orderAlerts.push({
          location: loc,
          orderDetails
        });
      }
    });

    setResults(analysisResults);
    setMixedPowerAlerts(alerts);
    setMixedOrderAlerts(orderAlerts);
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

  const exportToExcel = async () => {
    if (results.length === 0) return;
    
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('库位统计');

    // 定义列
    worksheet.columns = [
      { header: '库位名称', key: 'location', width: 20 },
      { header: '箱号数量', key: 'count', width: 15 },
      { header: '可入库数量 (标准18)', key: 'available', width: 25 }
    ];

    // 添加数据并设置样式
    results.forEach((row) => {
      const excelRow = worksheet.addRow({
        location: row['库位名称'],
        count: row['箱号数量'],
        available: row['可入库数量']
      });

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

    //表头样式
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
          boxes: boxes.join(', ')
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

  const exportOrdersToExcel = async () => {
    if (mixedOrderAlerts.length === 0) return;

    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('工单混载明细');

    worksheet.columns = [
      { header: '库位名称', key: 'location', width: 20 },
      { header: '工单号', key: 'order', width: 25 },
      { header: '包含箱号', key: 'boxes', width: 80 }
    ];

    mixedOrderAlerts.forEach((alert) => {
      Object.entries(alert.orderDetails).forEach(([order, boxes]) => {
        const excelRow = worksheet.addRow({
          location: alert.location,
          order: order,
          boxes: boxes.join(', ')
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
    saveAs(blob, `工单混载预警报告_${new Date().toISOString().split('T')[0]}.xlsx`);
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
    setMixedOrderAlerts([]);
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
            Excel 上传
          </button>
          
          <div className={cn(
            "w-full flex items-center gap-3 px-6 py-3 text-sm font-medium transition-colors",
            data.length ? "text-blue-400 font-bold" : "text-slate-700 opacity-50 cursor-not-allowed"
          )}>
            <TableIcon className="w-4 h-4" />
            统计结果分析
          </div>
          
          <div className="px-6 py-2 text-[10px] font-bold text-slate-500 uppercase tracking-widest mt-6 mb-2 opacity-50">工具</div>
          <div className="flex items-center gap-3 px-6 py-3 text-sm font-medium text-slate-500">
            <Hash className="w-4 h-4" />
            批量处理
          </div>
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
            <span className="px-2 py-0.5 bg-slate-100 text-[10px] text-slate-400 rounded-full font-bold uppercase tracking-wider">v1.2.0</span>
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
              {!data.length ? (
                <motion.div 
                  initial={{ opacity: 0, y: 10 }}
                  animate={{ opacity: 1, y: 0 }}
                  onDragOver={(e) => e.preventDefault()}
                  onDrop={handleDrop}
                  className="bg-white p-12 rounded-3xl border-2 border-dashed border-slate-200 flex flex-col items-center justify-center text-center relative group hover:border-blue-500 hover:bg-blue-50/10 transition-all duration-300 shadow-sm"
                >
                  <input 
                    type="file" 
                    accept=".xlsx,.xls" 
                    onChange={handleFileUpload}
                    className="absolute inset-0 opacity-0 cursor-pointer z-10"
                  />
                  <div className="w-20 h-20 bg-blue-50 rounded-3xl flex items-center justify-center mb-6 text-blue-600 group-hover:scale-110 group-hover:rotate-3 transition-transform duration-500">
                    <Upload className="w-10 h-10" />
                  </div>
                  <h3 className="font-bold text-xl mb-2 text-slate-900 group-hover:text-blue-600 transition-colors">拖拽库存文件到此处</h3>
                  <p className="text-sm text-slate-400 mb-8 max-w-[200px]">支持 Excel 格式 (.xlsx, .xls)</p>
                  <div className="px-6 py-3 bg-slate-900 text-white rounded-xl text-sm font-bold shadow-xl shadow-slate-200 transition-transform active:scale-95 group-hover:bg-blue-600">
                    选择本地文件
                  </div>
                  
                  {isProcessing && (
                    <div className="absolute inset-0 bg-white/95 backdrop-blur-sm rounded-3xl flex flex-col items-center justify-center z-20">
                      <Loader2 className="w-10 h-10 text-blue-600 animate-spin mb-4" />
                      <p className="text-sm font-bold text-slate-800 tracking-widest uppercase">正在解析...</p>
                    </div>
                  )}
                </motion.div>
              ) : (
                <motion.div 
                  initial={{ opacity: 0, scale: 0.95 }}
                  animate={{ opacity: 1, scale: 1 }}
                  className="bg-slate-900 rounded-3xl p-8 text-white shadow-2xl shadow-slate-900/10"
                >
                  <div className="flex items-center justify-between mb-8">
                    <h3 className="text-xs font-bold uppercase tracking-widest text-slate-500">统计摘要</h3>
                    <div className="flex items-center gap-1.5 px-2 py-1 bg-emerald-500/10 text-emerald-400 rounded-full border border-emerald-500/20">
                      <div className="w-1.5 h-1.5 rounded-full bg-emerald-500 animate-pulse" />
                      <span className="text-[10px] font-bold">同步中</span>
                    </div>
                  </div>
                  
                  <div className="space-y-6">
                    <div className="space-y-2">
                      <p className="text-[10px] text-slate-500 font-bold uppercase tracking-widest">总数据吞吐</p>
                      <div className="flex items-end gap-2">
                        <span className="text-4xl font-bold font-mono leading-none tracking-tighter">{data.length.toLocaleString()}</span>
                        <span className="text-xs text-slate-500 font-medium pb-1.5">行原始行</span>
                      </div>
                    </div>

                    <div className="space-y-2">
                      <p className="text-[10px] text-slate-500 font-bold uppercase tracking-widest">已识别库位</p>
                      <div className="flex items-end gap-2">
                        <span className="text-4xl font-bold font-mono leading-none tracking-tighter text-blue-400">{results.length}</span>
                        <span className="text-xs text-slate-500 font-medium pb-1.5">有效库区</span>
                      </div>
                    </div>

                    <div className="pt-6 space-y-3">
                      <button 
                        onClick={exportToExcel}
                        className="w-full flex items-center justify-center gap-3 bg-emerald-600 hover:bg-emerald-500 text-white font-bold py-4 rounded-2xl transition-all shadow-xl shadow-emerald-900/30 active:scale-[0.98]"
                      >
                        <Download className="w-5 h-5" />
                        导出库位统计表
                      </button>
                      
                      <button 
                        onClick={reset}
                        className="w-full text-slate-500 hover:text-slate-300 text-[10px] font-bold uppercase tracking-[0.2em] pt-4 transition-colors block text-center"
                      >
                        清空并重置环境
                      </button>
                    </div>
                  </div>
                </motion.div>
              )}

              <div className="bg-white p-7 rounded-3xl border border-slate-200 shadow-sm relative overflow-hidden group">
                <div className="absolute top-0 right-0 w-24 h-24 bg-blue-50 rounded-full -mr-12 -mt-12 group-hover:scale-110 transition-transform duration-700 opacity-60" />
                <h3 className="text-xs font-bold uppercase tracking-widest text-slate-400 mb-6 flex items-center gap-2">
                  <MapPin className="w-3 h-3" />
                  结构化字段要求
                </h3>
                <div className="space-y-4">
                  <div className="flex flex-col gap-1 p-4 bg-slate-50 rounded-2xl border border-slate-100 group-hover:bg-white transition-colors">
                    <span className="text-xs font-bold text-slate-900">库位名称</span>
                    <span className="text-[10px] text-slate-400 uppercase tracking-tighter">必填 • 字符串格式</span>
                  </div>
                  <div className="flex flex-col gap-1 p-4 bg-slate-50 rounded-2xl border border-slate-100 group-hover:bg-white transition-colors">
                    <span className="text-xs font-bold text-slate-900">箱号</span>
                    <span className="text-[10px] text-slate-400 uppercase tracking-tighter">必填 • 唯一标识符</span>
                  </div>
                  <p className="text-[10px] text-slate-400 leading-relaxed font-medium px-1">
                    系统将按每个<b>库位</b>作为主键，统计其所包含的独立不重复<b>箱号</b>之和。
                  </p>
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

                  {/* Mixed Order Alerts Summary */}
                  <motion.div 
                    initial={{ opacity: 0, y: 10 }}
                    animate={{ opacity: 1, y: 0 }}
                    transition={{ delay: 0.3 }}
                    className={cn(
                      "rounded-3xl p-6 border shadow-sm",
                      mixedOrderAlerts.length > 0 ? "bg-blue-50 border-blue-100" : "bg-white border-slate-200"
                    )}
                  >
                    <div className="flex items-center justify-between mb-4">
                      <div className="flex items-center gap-2">
                        <h4 className="text-sm font-bold flex items-center gap-2 text-nowrap">
                          <History className={cn("w-4 h-4", mixedOrderAlerts.length > 0 ? "text-blue-500" : "text-slate-400")} />
                          工单混载监测
                        </h4>
                        {mixedOrderAlerts.length > 0 && (
                          <button 
                            onClick={exportOrdersToExcel}
                            className="p-1.5 hover:bg-blue-100 text-blue-600 rounded-lg transition-colors group/btn"
                            title="导出工单混载"
                          >
                            <Download className="w-3.5 h-3.5" />
                          </button>
                        )}
                      </div>
                      <span className={cn("text-xl font-black", mixedOrderAlerts.length > 0 ? "text-blue-600" : "text-slate-400")}>
                        {mixedOrderAlerts.length}
                      </span>
                    </div>
                    <div className="max-h-24 overflow-y-auto pr-2 scrollbar-thin">
                      {mixedOrderAlerts.length > 0 ? (
                        <div className="space-y-3">
                          {mixedOrderAlerts.map((alert, i) => (
                            <div key={i} className="flex flex-col p-3 bg-white rounded-xl border border-blue-200 text-[10px] shadow-sm">
                              <div className="font-bold text-slate-800 mb-2 flex items-center gap-1">
                                <Package className="w-3 h-3 text-blue-400" />
                                {alert.location}
                              </div>
                              <div className="space-y-1">
                                {Object.keys(alert.orderDetails).map(o => (
                                  <div key={o} className="flex items-center justify-between text-slate-500">
                                    <span className="truncate max-w-[100px]" title={o}>{o}</span>
                                    <span className="font-mono text-blue-600 bg-blue-50 px-1.5 rounded">{alert.orderDetails[o].length}箱</span>
                                  </div>
                                ))}
                              </div>
                            </div>
                          ))}
                        </div>
                      ) : (
                        <p className="text-slate-400 text-xs italic">库位工单单一，未发现混载</p>
                      )}
                    </div>
                  </motion.div>
                </div>
              )}

              {/* Main Detail Table */}
              <div className="bg-white rounded-3xl border border-slate-200 shadow-sm overflow-hidden flex flex-col flex-1">
                <div className="p-8 border-b border-slate-100 flex items-center justify-between bg-white z-20">
                  <div>
                    <h3 className="font-bold text-xl text-slate-900 tracking-tight">分析明细看板</h3>
                    <p className="text-xs text-slate-400 font-medium">分组统计库位上的箱号总量</p>
                  </div>
                  {results.length > 0 && (
                    <div className="hidden sm:flex items-center gap-4">
                      <div className="text-right">
                        <p className="text-[10px] font-bold text-slate-400 uppercase tracking-widest">记录数</p>
                        <p className="text-sm font-bold text-slate-800">{results.length}</p>
                      </div>
                      <div className="w-px h-8 bg-slate-100" />
                      <button 
                        onClick={exportToExcel}
                        className="p-2.5 bg-slate-50 hover:bg-slate-100 text-slate-600 rounded-xl border border-slate-200 transition-colors"
                      >
                        <Download className="w-4 h-4" />
                      </button>
                    </div>
                  )}
                </div>

                <div className="flex-1 overflow-auto">
                  <table className="w-full border-collapse">
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
                                  "w-2 h-2 rounded-full scale-50 group-hover:scale-100 transition-all",
                                  row['箱号数量'] > 18 ? "bg-red-500 shadow-[0_0_8px_rgba(239,68,68,0.5)]" : 
                                  row['箱号数量'] < 18 ? "bg-emerald-500 shadow-[0_0_8px_rgba(16,185,129,0.5)]" : 
                                  "bg-slate-900 shadow-[0_0_8px_rgba(15,23,42,0.3)]"
                                )} />
                                <span className={cn(
                                  "text-sm font-bold transition-colors",
                                  row['箱号数量'] > 18 ? "text-red-600" : 
                                  row['箱号数量'] < 18 ? "text-emerald-700" : 
                                  "text-slate-900"
                                )}>
                                  {row['库位名称']}
                                </span>
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
                          <td colSpan={3} className="px-8 py-32 text-center">
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

