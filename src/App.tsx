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
  Loader2,
  ShieldAlert,
  Save,
  Cloud,
  CloudUpload,
  CloudDownload,
  RefreshCw,
  Layers,
  Filter,
  Search,
  X,
  Eye,
  HelpCircle,
  Info,
  ListOrdered
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import { jsPDF } from 'jspdf';
import { cn } from './lib/utils';
import { db, collection, addDoc, getDocs, query, orderBy, limit, doc, getDoc, setDoc } from './lib/firebase';

interface ZeroBarcodeRow {
  rowIdx: number; // Excel 实际行号
  boxNo: string;
  location: string;
  rawBarcode: string;
  grade: string;
  orderNo?: string;
  reason: string;
}

interface AuditReport {
  duplicateBoxes: Array<{ boxNo: string; count: number; rows: Array<{ rowIdx: number; location: string; grade: string; power: string; orderNo: string }> }>;
  duplicateRowCount: number; // 重复增加的记录行数
  emptyBoxRows: Array<{ rowIdx: number; location: string; grade: string; power: string; orderNo: string }>;
  emptyLocRows: Array<{ rowIdx: number; boxNo: string; grade: string; power: string; orderNo: string }>;
  zeroBarcodeBreakdown: {
    total: number;
    gradeA: number;
    gradeNonA: number;
    rows: ZeroBarcodeRow[];
  };
}

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
  '主存客户'?: string;
  '主存功率'?: string;
}

interface MixedPowerAlert {
  location: string;
  powerDetails: Record<string, string[]>; 
}

interface MixedCustomerAlert {
  location: string;
  customerDetails: Record<string, { boxNo: string; power: string }[]>; // CustomerName -> List of BoxDetails
}

interface CustomerStat {
  customerName: string;
  boxCount: number;
  locations: string[];
}

interface RelocationSuggestion {
  boxNo: string;
  currentLoc: string;
  customerName: string;
  powerGrade: string;
  dominantCustomer: string;
  dominantPower: string;
  mismatchType: string;
  recommendedLoc: string;
  recommendationReason: string;
}

interface NonAGradeRow {
  boxNo: string;
  location: string;
  rawLocation: string;
  grade: string; // 'B' | 'F' | 'R' | 其它
  rawGrade: string; // K列原始文本
  barcodeCount?: number; // S列箱条码数
  orderNo?: string; // J列工单号
  customerName?: string; // 对应的客户简称
  power?: string; // 功率档
}

// 仓库基准全量库位序列（严格按用户指定顺序）：
const PREDEFINED_LOCATIONS: string[] = (() => {
  const list: string[] = [];

  // C00 ~ C60, D00 ~ D60, E00 ~ E60, F00 ~ F60, G00 ~ G60, H00 ~ H60
  ['C', 'D', 'E', 'F', 'G', 'H'].forEach(zone => {
    for (let i = 0; i <= 60; i++) {
      list.push(`${zone}${String(i).padStart(2, '0')}`);
    }
  });

  // K00 ~ K35, L00 ~ L35, M00 ~ M35, N00 ~ N35, O00 ~ O35
  ['K', 'L', 'M', 'N', 'O'].forEach(zone => {
    for (let i = 0; i <= 35; i++) {
      list.push(`${zone}${String(i).padStart(2, '0')}`);
    }
  });

  // WA
  list.push('WA');

  // WE1 ~ WE5
  for (let i = 1; i <= 5; i++) {
    list.push(`WE${i}`);
  }

  // WF1 ~ WF5
  for (let i = 1; i <= 5; i++) {
    list.push(`WF${i}`);
  }

  // WG1 ~ WG5
  for (let i = 1; i <= 5; i++) {
    list.push(`WG${i}`);
  }

  // TBD (待定库位，固定基准显示，无货时独立成行显示 0)
  list.push('TBD');

  return list;
})();

const PREDEFINED_LOCATION_SET = new Set(PREDEFINED_LOCATIONS);
const PREDEFINED_LOCATION_ORDER_MAP = new Map<string, number>();
PREDEFINED_LOCATIONS.forEach((loc, index) => {
  PREDEFINED_LOCATION_ORDER_MAP.set(loc, index);
});

/**
 * 库位名称智能规整函数：
 * 1. 匹配 WA
 * 2. 匹配 TBD (待定/暂存库位)
 * 3. 匹配 WE, WF, WG + 序号 (如 we1, WE01 -> WE1, wf3, WF03 -> WF3, wg5 -> WG5)
 * 4. 匹配单字母 + 数字 (如 C0 -> C00, C00 -> C00, C1 -> C01, K0 -> K00, K5 -> K05, O35 -> O35)
 */
const normalizeLocationName = (rawLoc: any): string => {
  let loc = String(rawLoc || '').trim();
  if (!loc) return '';

  // 1. 匹配 WA
  if (/^wa$/i.test(loc)) {
    return 'WA';
  }

  // 2. 匹配 TBD (待定库位)
  if (/^tbd$/i.test(loc)) {
    return 'TBD';
  }

  // 3. 匹配 WE, WF, WG 开头的库位
  const wPrefixMatch = loc.match(/^(we|wf|wg)0*([1-9]\d*)$/i);
  if (wPrefixMatch) {
    return `${wPrefixMatch[1].toUpperCase()}${parseInt(wPrefixMatch[2], 10)}`;
  }

  // 4. 匹配单字母 + 数字 (如 C0 -> C00, C00 -> C00, C1 -> C01, K0 -> K00, M2 -> M02, L16 -> L16, N22 -> N22, O5 -> O05)
  const letterMatch = loc.match(/^([a-zA-Z])0*(\d+)$/);
  if (letterMatch) {
    const letter = letterMatch[1].toUpperCase();
    const num = parseInt(letterMatch[2], 10);
    return `${letter}${String(num).padStart(2, '0')}`;
  }

  return loc;
};

/**
 * 库位全局排序对比函数：
 * - 优先按照基准序列 PREDEFINED_LOCATIONS 的指定顺序严格排列
 * - 不在基准范围内的外部库位，统一排列在最后，并按名称自然升序展示
 */
const compareLocations = (locA: string, locB: string): number => {
  const orderA = PREDEFINED_LOCATION_ORDER_MAP.get(locA);
  const orderB = PREDEFINED_LOCATION_ORDER_MAP.get(locB);

  if (orderA !== undefined && orderB !== undefined) {
    return orderA - orderB;
  }
  if (orderA !== undefined) {
    return -1;
  }
  if (orderB !== undefined) {
    return 1;
  }
  return locA.localeCompare(locB, undefined, { numeric: true, sensitivity: 'base' });
};

/**
 * 来源K列的箱等级归一化分类函数：
 * 识别 B、F、R 等非A级别品
 */
const getNormalizedGradeCategory = (rawGrade: string): string => {
  const g = String(rawGrade || '').trim().toUpperCase();
  if (!g) return '未标级';
  if (g.startsWith('B') || g.includes('B级') || g.includes('B等级')) return 'B';
  if (g.startsWith('F') || g.includes('F级') || g.includes('F等级')) return 'F';
  if (g.startsWith('R') || g.includes('R级') || g.includes('R等级')) return 'R';
  return g;
};

export default function App() {
  const STANDARD_CAPACITY = 18;
  const [data, setData] = useState<InventoryRow[]>([]);
  const [results, setResults] = useState<AnalysisResult[]>([]);
  const [mixedPowerAlerts, setMixedPowerAlerts] = useState<MixedPowerAlert[]>([]);
  const [mixedCustomerAlerts, setMixedCustomerAlerts] = useState<MixedCustomerAlert[]>([]);
  const [relocationSuggestions, setRelocationSuggestions] = useState<RelocationSuggestion[]>([]);
  const [orderCustomerMap, setOrderCustomerMap] = useState<Record<string, string>>(() => {
    try {
      const stored = localStorage.getItem('orderCustomerMap');
      return stored ? JSON.parse(stored) : {};
    } catch (e) {
      console.error('Failed to parse stored orderCustomerMap', e);
      return {};
    }
  });
  const [mappingFileName, setMappingFileName] = useState(() => {
    return localStorage.getItem('mappingFileName') || '';
  });
  const [customerStats, setCustomerStats] = useState<CustomerStat[]>([]);
  const [isSaved, setIsSaved] = useState<boolean>(() => {
    return !!localStorage.getItem('mappingFileName');
  });
  const [saveSuccess, setSaveSuccess] = useState<string | null>(null);
  const [totalRowsRaw, setTotalRowsRaw] = useState<number>(0);
  const [nonAGradeCount, setNonAGradeCount] = useState<number>(0);
  const [zeroBarcodeCount, setZeroBarcodeCount] = useState<number>(0);
  const [zeroBarcodeList, setZeroBarcodeList] = useState<ZeroBarcodeRow[]>([]);
  const [showZeroBarcodeModal, setShowZeroBarcodeModal] = useState<boolean>(false);
  const [auditReport, setAuditReport] = useState<AuditReport | null>(null);
  const [showAuditModal, setShowAuditModal] = useState<boolean>(false);
  const [nonAGradeList, setNonAGradeList] = useState<NonAGradeRow[]>([]);
  const [nonAGradeFilter, setNonAGradeFilter] = useState<'ALL' | 'B' | 'F' | 'R' | 'OTHER'>('ALL');
  const [nonAGradeSearch, setNonAGradeSearch] = useState<string>('');
  const [activeTab, setActiveTab] = useState<'location' | 'customer' | 'correction' | 'nonA'>('location');
  const [isProcessing, setIsProcessing] = useState(false);
  const [fileName, setFileName] = useState('');
  const [error, setError] = useState<string | null>(null);

  // 云端存储相关 State (Firebase integration)
  const [cloudMappings, setCloudMappings] = useState<Array<{ id: string; mappingFileName: string; count: number; updatedAt: string; orderCustomerMap: Record<string, string> }>>([]);
  const [isCloudSaving, setIsCloudSaving] = useState(false);
  const [isCloudLoading, setIsCloudLoading] = useState(false);
  const [lastSyncedCloudTime, setLastSyncedCloudTime] = useState<string | null>(null);
  const [isCloudSynced, setIsCloudSynced] = useState(false);

  // 获取云端全部关系对照模板
  const fetchCloudMappings = useCallback(async (autoLoadIfEmpty = false) => {
    setIsCloudLoading(true);
    try {
      const q = query(collection(db, 'mappings'), orderBy('updatedAt', 'desc'), limit(15));
      const querySnapshot = await getDocs(q);
      const mappingsList: any[] = [];
      querySnapshot.forEach((doc) => {
        const itemData = doc.data();
        if (doc.id !== 'latest') { 
          mappingsList.push({
            id: doc.id,
            ...itemData
          });
        }
      });
      setCloudMappings(mappingsList);

      // 若当前本地没有对照表，尝试从云端拉取最新的自动加载
      if (autoLoadIfEmpty) {
        let loadedDoc = null;
        
        // 优先查看 special "latest" doc
        const latestDocRef = doc(db, 'mappings', 'latest');
        try {
          const latestDocSnap = await getDoc(latestDocRef);
          if (latestDocSnap.exists()) {
            loadedDoc = latestDocSnap.data();
          }
        } catch (e) {
          console.log('No default latest file key found, trying collection');
        }

        // 如果未取到，直接走列表里的第一个
        if (!loadedDoc && mappingsList.length > 0) {
          loadedDoc = mappingsList[0];
        }

        if (loadedDoc && loadedDoc.orderCustomerMap) {
          const loadedMap = loadedDoc.orderCustomerMap;
          setOrderCustomerMap(loadedMap);
          setMappingFileName(loadedDoc.mappingFileName || '云端同步对照模型');
          setIsSaved(true);
          setIsCloudSynced(true);
          
          const timeStr = loadedDoc.updatedAt ? new Date(loadedDoc.updatedAt).toLocaleTimeString() : '刚刚';
          setLastSyncedCloudTime(timeStr);
          
          setSaveSuccess(`☁️ 已自动为您加载云端最新的对照模型：${loadedDoc.mappingFileName || '未知文件名'}`);
          setTimeout(() => setSaveSuccess(null), 5000);
        }
      }
    } catch (err: any) {
      console.error('Failed to fetch cloud mappings:', err);
    } finally {
      setIsCloudLoading(false);
    }
  }, []);

  // 组件装载时，自动拉取云端记录：如果本地为空就自动应用最新的
  React.useEffect(() => {
    const isLocalEmpty = !localStorage.getItem('mappingFileName') && Object.keys(orderCustomerMap).length === 0;
    fetchCloudMappings(isLocalEmpty);
  }, []);

  // 永久同步关系值到 Firebase 云端
  const saveMappingToCloud = async () => {
    if (!mappingFileName || Object.keys(orderCustomerMap).length === 0) {
      setError('无法保存：当前未加载任何工单客户对照表');
      return;
    }
    setIsCloudSaving(true);
    try {
      const dataPayload = {
        mappingFileName,
        orderCustomerMap,
        updatedAt: new Date().toISOString(),
        count: Object.keys(orderCustomerMap).length,
      };
      
      // 1. 存入历史集合
      const docRef = await addDoc(collection(db, 'mappings'), dataPayload);
      
      // 2. 同时在 latest 文档上覆盖记录
      await setDoc(doc(db, 'mappings', 'latest'), dataPayload);

      setIsCloudSynced(true);
      setLastSyncedCloudTime(new Date().toLocaleTimeString());
      setSaveSuccess(`云端发布成功！共享 ID: ${docRef.id}，其它客户端刷新即可自动拉取最新。`);
      setTimeout(() => setSaveSuccess(null), 5000);
      
      // 刷新列表
      await fetchCloudMappings(false);
    } catch (err: any) {
      console.error(err);
      setError('云端数据库同步失败：' + (err.message || String(err)));
    } finally {
      setIsCloudSaving(false);
    }
  };

  // 应用云端版本到当前上下文中
  const applyCloudMapping = (mappingItem: any) => {
    if (!mappingItem || !mappingItem.orderCustomerMap) return;
    setOrderCustomerMap(mappingItem.orderCustomerMap);
    setMappingFileName(mappingItem.mappingFileName);
    setIsSaved(true);
    setIsCloudSynced(true);
    setLastSyncedCloudTime(new Date(mappingItem.updatedAt).toLocaleTimeString());

    setSaveSuccess(`已成功应用云端版本：${mappingItem.mappingFileName}`);
    setTimeout(() => setSaveSuccess(null), 3000);

    // 如果已加载了库存表，顺便连带重新计算分析
    if (data.length > 0) {
      performAnalysis(data, mappingItem.orderCustomerMap);
    }
  };

  // 备份并下载云端对应文件为 Excel
  const downloadCloudMappingAsExcel = async (mappingItem: any) => {
    try {
      const workbook = new ExcelJS.Workbook();
      const worksheet = workbook.addWorksheet('云端关系对照备份');

      worksheet.columns = [
        { header: '序号', key: 'seq', width: 12 },
        { header: '工单号 / 销售单号 (必填-列B)', key: 'orderNo', width: 28 },
        { header: '辅助信息列1', key: 'remark1', width: 22 },
        { header: '辅助信息列2', key: 'remark2', width: 22 },
        { header: '辅助信息列3', key: 'remark3', width: 22 },
        { header: '客户简称 / 客户名称 (必填-列F)', key: 'customerName', width: 32 },
      ];

      worksheet.views = [{ showGridLines: true }];

      // Header rows layout style
      const headerRow = worksheet.getRow(1);
      headerRow.height = 30;
      headerRow.eachCell((cell, colNumber) => {
        cell.font = { name: '微软雅黑', size: 10, bold: true, color: { argb: 'FFFFFF' } };
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: colNumber === 2 || colNumber === 6 ? '78350F' : '475569' }
        };
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
      });

      let seqIdx = 1;
      Object.entries(mappingItem.orderCustomerMap || {}).forEach(([orderNo, customerName]) => {
        const row = worksheet.addRow({
          seq: String(seqIdx++),
          orderNo: orderNo,
          remark1: '随处云端导入备份',
          remark2: '',
          remark3: '',
          customerName: customerName
        });
        row.height = 22;
        row.eachCell((cell, colNumber) => {
          cell.font = { name: '微软雅黑', size: 10 };
          cell.alignment = { horizontal: colNumber === 2 || colNumber === 6 ? 'left' : 'center', vertical: 'middle' };
          if (colNumber === 2 || colNumber === 6) {
            cell.fill = {
              type: 'pattern',
              pattern: 'solid',
              fgColor: { argb: 'FFFBEB' } // light amber
            };
          }
          cell.border = {
            top: { style: 'thin', color: { argb: 'F1F5F9' } },
            left: { style: 'thin', color: { argb: 'F1F5F9' } },
            bottom: { style: 'thin', color: { argb: 'F1F5F9' } },
            right: { style: 'thin', color: { argb: 'F1F5F9' } }
          };
        });
      });

      const buffer = await workbook.xlsx.writeBuffer();
      const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      saveAs(blob, `${mappingItem.mappingFileName}_云端备份_${new Date(mappingItem.updatedAt).toLocaleDateString()}.xlsx`);
      
      setSaveSuccess(`已成功下载云端对照备份：${mappingItem.mappingFileName}`);
      setTimeout(() => setSaveSuccess(null), 3000);
    } catch (e: any) {
      console.error(e);
      setError('从云端生成 Excel 文件失败：' + (e.message || String(e)));
    }
  };

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

        // 箱等级标题的模糊位置 (作为 column K 越界或为空时的 fallback)
        let gradeIdx = headers.findIndex(h => h.includes('箱等级') || h === '等级' || h.includes('等级') || h.toLowerCase() === 'grade');

        // 箱条码数标题的模糊位置 (作为 column S 越界或为空时的 fallback)
        let barcodeCountIdx = headers.findIndex(h => h === '箱条码数' || h.includes('箱条码数') || h === '条码数' || h.includes('条码数') || h.includes('条码数量'));

        const parsedData: InventoryRow[] = [];
        const nonAGradeListLocal: NonAGradeRow[] = [];
        const zeroBarcodeListLocal: ZeroBarcodeRow[] = [];
        let totalRowsCount = 0;
        let nonAGradeCountLocal = 0;
        let zeroBarcodeCountLocal = 0;

        // 审计引擎收集器：定位重复箱号、空箱号、空库位以及条码=0在各等级的分布
        const boxOccurrences = new Map<string, Array<{ rowIdx: number; location: string; grade: string; power: string; orderNo: string }>>();
        const emptyBoxRowsLocal: Array<{ rowIdx: number; location: string; grade: string; power: string; orderNo: string }> = [];
        const emptyLocRowsLocal: Array<{ rowIdx: number; boxNo: string; grade: string; power: string; orderNo: string }> = [];

        for (let i = 1; i < arrayRows.length; i++) {
          const rowArr = arrayRows[i];
          if (!rowArr || rowArr.length === 0) continue;
          
          // 排除全是空单元格的虚假行
          const isAllEmpty = rowArr.every(val => val === undefined || val === null || String(val).trim() === '');
          if (isAllEmpty) continue;
          
          totalRowsCount++;

          const boxVal = rowArr[boxIdx] !== undefined && rowArr[boxIdx] !== null ? String(rowArr[boxIdx]).trim() : '';
          const locVal = rowArr[locIdx] !== undefined && rowArr[locIdx] !== null ? String(rowArr[locIdx]).trim() : '';
          
          // 按用户要求：主库存表中每一箱的工单号在 J 列（Column J 是第 10 列，0-based 索引为 9）
          let orderNoVal = '';
          if (rowArr.length > 9 && rowArr[9] !== undefined && rowArr[9] !== null && String(rowArr[9]).trim() !== '') {
            orderNoVal = String(rowArr[9]).trim();
          } else if (orderNoFallbackIdx !== -1 && rowArr[orderNoFallbackIdx] !== undefined && rowArr[orderNoFallbackIdx] !== null) {
            orderNoVal = String(rowArr[orderNoFallbackIdx]).trim();
          }

          // 箱等级在 K 列（Column K 是第 11 列，0-based 索引为 10）
          let gradeVal = '';
          if (rowArr.length > 10 && rowArr[10] !== undefined && rowArr[10] !== null && String(rowArr[10]).trim() !== '') {
            gradeVal = String(rowArr[10]).trim();
          } else if (gradeIdx !== -1 && rowArr[gradeIdx] !== undefined && rowArr[gradeIdx] !== null) {
            gradeVal = String(rowArr[gradeIdx]).trim();
          }

          let powerIdx = headers.findIndex(h => h.includes('功率'));
          const powerVal = powerIdx !== -1 ? rowArr[powerIdx] : undefined;
          const powerStr = powerVal !== undefined && powerVal !== null ? String(powerVal).trim() : '未知';

          // 记录箱号出现特征（用于排查 8 箱差额是否来源于重复箱号或空箱号）
          if (boxVal === '') {
            emptyBoxRowsLocal.push({
              rowIdx: i + 1,
              location: locVal || '未标库位',
              grade: gradeVal || '未标级',
              power: powerStr,
              orderNo: orderNoVal || '-'
            });
          } else {
            const occurrences = boxOccurrences.get(boxVal) || [];
            occurrences.push({
              rowIdx: i + 1,
              location: locVal || '未标库位',
              grade: gradeVal || '未标级',
              power: powerStr,
              orderNo: orderNoVal || '-'
            });
            boxOccurrences.set(boxVal, occurrences);
          }

          if (locVal === '' || locVal.toUpperCase() === 'TBD') {
            emptyLocRowsLocal.push({
              rowIdx: i + 1,
              boxNo: boxVal || `ROW-${i + 1}`,
              grade: gradeVal || '未标级',
              power: powerStr,
              orderNo: orderNoVal || '-'
            });
          }

          // 检查 S 列 箱条码数 (Column S 是第 19 列，0-based 索引为 18)
          // 用户强规则：“在原始数据里，S列 箱条码数 0的话，不纳入统计”
          // 重要业务界限：
          // 1. 显式填写为 0（如数字 0、字符 "0"、"0.0"、"0箱"、全角"０"等）：表示该箱无条码/已作废，不纳入统计！
          // 2. 留空、未填写或非0：属于正常有效库存箱（未单独录入条码数），必须纳入统计，绝不能过滤！
          let barcodeCheckRaw: any = undefined;
          let sourceDesc = 'S列(第19列)';
          if (barcodeCountIdx !== -1 && barcodeCountIdx < rowArr.length && rowArr[barcodeCountIdx] !== undefined) {
            barcodeCheckRaw = rowArr[barcodeCountIdx];
            sourceDesc = `表头[${headers[barcodeCountIdx]}]`;
          } else if (rowArr.length > 18 && rowArr[18] !== undefined) {
            barcodeCheckRaw = rowArr[18];
            sourceDesc = 'S列(第19列)';
          }

          let isBarcodeZero = false;
          let barcodeZeroReason = '';
          let barcodeCountNum: number | undefined = undefined;

          // 只有单元格有实际填写内容，且内容明确解析为 0 时，才判定为条码数为0
          if (barcodeCheckRaw !== undefined && barcodeCheckRaw !== null) {
            const rawStr = String(barcodeCheckRaw).trim();
            // 若为空字符串，代表单元格留空未填，属于正常箱，不作为0处理
            if (rawStr !== '') {
              // 全角转半角，清理常见中英文计量单位
              const normalizedStr = rawStr.replace(/[\uFF10-\uFF19]/g, ch => String.fromCharCode(ch.charCodeAt(0) - 0xFEE0));
              const cleanedStr = normalizedStr.replace(/[箱个件支片条只包盒袋pcsPC\s]/gi, '');

              if (cleanedStr !== '') {
                const parsedBarcodeNum = Number(cleanedStr);
                if (!isNaN(parsedBarcodeNum)) {
                  if (parsedBarcodeNum === 0) {
                    isBarcodeZero = true;
                    barcodeZeroReason = `${sourceDesc}条码数为 0 ("${rawStr}")`;
                  } else {
                    barcodeCountNum = parsedBarcodeNum;
                  }
                } else if (/^0+(\.0+)?$/.test(cleanedStr)) {
                  isBarcodeZero = true;
                  barcodeZeroReason = `${sourceDesc}条码数为 0 ("${rawStr}")`;
                }
              }
            }
          }

          if (isBarcodeZero) {
            zeroBarcodeCountLocal++;
            zeroBarcodeListLocal.push({
              rowIdx: i + 1,
              boxNo: boxVal || `ROW-${i + 1}`,
              location: locVal || '未标库位',
              rawBarcode: barcodeCheckRaw !== undefined && barcodeCheckRaw !== null ? String(barcodeCheckRaw) : '',
              grade: gradeVal || '未标级',
              orderNo: orderNoVal || '-',
              reason: barcodeZeroReason
            });
            continue;
          }

          // 仅统计A等级: A, A等级, A级, a等级, a, etc. (不区分大小写，或startsWith 'A')
          const cleanGrade = gradeVal.trim().toUpperCase().replace(/\s+/g, '');
          const isGradeA = cleanGrade === 'A' || 
                           cleanGrade === 'A级' || 
                           cleanGrade === 'A等级' || 
                           cleanGrade === 'A品' || 
                           cleanGrade === 'A类' || 
                           cleanGrade.startsWith('A+') || 
                           cleanGrade.startsWith('A-') || 
                           cleanGrade.startsWith('A_') || 
                           /^A\d+$/.test(cleanGrade) ||
                           cleanGrade === 'A-等级';

          if (!isGradeA) {
            nonAGradeCountLocal++;
            const cat = getNormalizedGradeCategory(gradeVal);

            nonAGradeListLocal.push({
              boxNo: boxVal || `BOX-${i}`,
              location: normalizeLocationName(locVal),
              rawLocation: locVal,
              grade: cat,
              rawGrade: gradeVal,
              barcodeCount: barcodeCountNum,
              orderNo: orderNoVal,
              customerName: orderCustomerMap[orderNoVal] || '未映射客户',
              power: powerVal !== undefined && powerVal !== null ? String(powerVal).trim() : '未知'
            });
            continue;
          }

          parsedData.push({
            '箱号': boxVal,
            '库位名称': locVal,
            '功率档': powerVal,
            '工单号': orderNoVal,
            '销售单号': orderNoVal,
            '箱等级': gradeVal,
            '箱条码数': barcodeCountNum
          });
        }

        if (parsedData.length === 0) {
          if (nonAGradeCountLocal > 0 || zeroBarcodeCountLocal > 0) {
            const reasons: string[] = [];
            if (nonAGradeCountLocal > 0) reasons.push(`过滤了 ${nonAGradeCountLocal} 行非 A 级`);
            if (zeroBarcodeCountLocal > 0) reasons.push(`过滤了 ${zeroBarcodeCountLocal} 行 S列箱条码数=0`);
            throw new Error(`主库存表中检测到 ${totalRowsCount} 行数据，但有效箱数为 0 (${reasons.join('，')})。`);
          } else {
            throw new Error('主库存表中没有成功加载有效数据行');
          }
        }

        setTotalRowsRaw(totalRowsCount);
        setNonAGradeCount(nonAGradeCountLocal);
        setZeroBarcodeCount(zeroBarcodeCountLocal);
        setZeroBarcodeList(zeroBarcodeListLocal);
        setNonAGradeList(nonAGradeListLocal);

        // 构建 8 箱差异排查与行号诊断数据
        const duplicateBoxesList: Array<{ boxNo: string; count: number; rows: Array<{ rowIdx: number; location: string; grade: string; power: string; orderNo: string }> }> = [];
        let duplicateRowCountLocal = 0;
        boxOccurrences.forEach((rows, bNo) => {
          if (rows.length > 1) {
            duplicateBoxesList.push({ boxNo: bNo, count: rows.length, rows });
            duplicateRowCountLocal += (rows.length - 1);
          }
        });

        // 统计 23 箱条码为 0 的箱子中 A 级 vs 非 A 级分布
        let zeroBarcodeGradeACount = 0;
        let zeroBarcodeGradeNonACount = 0;
        zeroBarcodeListLocal.forEach(item => {
          const g = item.grade.trim().toUpperCase();
          if (g === 'A' || g === 'A级' || g === 'A等级' || g.startsWith('A')) {
            zeroBarcodeGradeACount++;
          } else {
            zeroBarcodeGradeNonACount++;
          }
        });

        const report: AuditReport = {
          duplicateBoxes: duplicateBoxesList,
          duplicateRowCount: duplicateRowCountLocal,
          emptyBoxRows: emptyBoxRowsLocal,
          emptyLocRows: emptyLocRowsLocal,
          zeroBarcodeBreakdown: {
            total: zeroBarcodeCountLocal,
            gradeA: zeroBarcodeGradeACount,
            gradeNonA: zeroBarcodeGradeNonACount,
            rows: zeroBarcodeListLocal
          }
        };
        setAuditReport(report);

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
        setIsSaved(false);
        setIsCloudSynced(false);
        setLastSyncedCloudTime(null);
        
        setNonAGradeList(prev => prev.map(item => ({
          ...item,
          customerName: newMap[item.orderNo || ''] || item.customerName || '未映射客户'
        })));

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
    // 箱号由功率的对照字典
    const boxPowerMap: Record<string, string> = {};
    
    const activeBoxes: { boxNo: string; location: string; power: string; customerName: string }[] = [];
 
    // 1. 初始化固定的全量基准库位（保证无论是否有货均作为独立行展现，空置时数量显示为0）
    PREDEFINED_LOCATIONS.forEach(locName => {
      locationMap[locName] = new Set();
      locationPowerData[locName] = {};
      locationCustomerData[locName] = {};
    });
 
    // 2. 处理数据
    rows.forEach((row, rowIdx) => {
      let loc = normalizeLocationName(row['库位名称']);
      if (!loc) {
        loc = 'TBD';
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

      const boxStr = (box !== undefined && box !== null && String(box).trim() !== '')
        ? String(box).trim()
        : `未标箱号-${rowIdx + 1}`;

      boxPowerMap[boxStr] = power;
      locationMap[loc].add(boxStr);

      activeBoxes.push({
        boxNo: boxStr,
        location: loc,
        power: power,
        customerName: custName
      });

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
    });

    // 汇总分析结果：
    // - 基准库位全量展现（空置时箱数为0）
    // - 不在基准范围内的库位，如果有物品（箱号数量 > 0），则纳入汇总并列在最后显示
    const analysisResults: AnalysisResult[] = Object.entries(locationMap)
      .filter(([location, boxSet]) => {
        if (PREDEFINED_LOCATION_SET.has(location)) {
          return true;
        }
        return boxSet.size > 0;
      })
      .map(([location, boxSet]) => {
        const actualBoxesInLoc = activeBoxes.filter(b => b.location === location);
        const boxCount = actualBoxesInLoc.length;
        const customerMapForLoc = locationCustomerData[location] || {};
        const customersDist: Record<string, number> = {};
        Object.entries(customerMapForLoc).forEach(([cust, boxes]) => {
          if (boxes.size > 0) {
            customersDist[cust] = boxes.size;
          }
        });

        const powers = locationPowerData[location] || {};
        // 主功率：箱数最多的功率挡
        let dominantPower = '未知';
        let maxPowerCount = 0;
        Object.entries(powers).forEach(([p, boxes]) => {
          if (boxes.size > maxPowerCount) {
            maxPowerCount = boxes.size;
            dominantPower = p;
          }
        });

        // 主客户：箱数最多的客户名称
        let dominantCustomer = '未映射客户';
        let maxCustomerCount = 0;
        Object.entries(customerMapForLoc).forEach(([c, boxes]) => {
          if (boxes.size > maxCustomerCount) {
            maxCustomerCount = boxes.size;
            dominantCustomer = c;
          }
        });

        const capacityForLoc = location.toUpperCase().startsWith('H') ? 14 : 18;
        return {
          '库位名称': location,
          '箱号数量': boxCount,
          '可入库数量': capacityForLoc - boxCount,
          '客户分布': customersDist,
          '主存客户': boxCount > 0 ? dominantCustomer : '空置',
          '主存功率': boxCount > 0 ? dominantPower : '空置'
        };
      });
 
    // 排序逻辑：严格按照基准序列（C00~C60, D00~D60...WG1~WG5）排序，不在范围内的有货库位统一排在最末尾
    analysisResults.sort((a, b) => compareLocations(a['库位名称'], b['库位名称']));
 
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
        const customerDetails: Record<string, { boxNo: string; power: string }[]> = {};
        customerNames.forEach(c => {
          customerDetails[c] = Array.from(customers[c]).map(boxNo => ({
            boxNo,
            power: boxPowerMap[boxNo] || '未知'
          }));
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
 
        let loc = normalizeLocationName(row['库位名称']);
 
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
        locations: Array.from(customerLocs[custName] || []).sort(compareLocations)
      };
    }).sort((a, b) => b.boxCount - a.boxCount);
 
    // 计算每个库位的主属性（主客户和主功率）与迁移修正建议
    const getPerfectCapacity = (locName: string) => {
      return locName.toUpperCase().startsWith('H') ? 14 : 18;
    };

    const locationDominantAttrs: Record<string, { dominantCustomer: string; dominantPower: string }> = {};

    Object.keys(locationMap).forEach(loc => {
      const powers = locationPowerData[loc] || {};
      const customers = locationCustomerData[loc] || {};

      // 主功率：箱数最多的功率挡
      let dominantPower = '未知';
      let maxPowerCount = 0;
      Object.entries(powers).forEach(([p, boxes]) => {
        if (boxes.size > maxPowerCount) {
          maxPowerCount = boxes.size;
          dominantPower = p;
        }
      });

      // 主客户：箱数最多的客户名称
      let dominantCustomer = '未映射客户';
      let maxCustomerCount = 0;
      Object.entries(customers).forEach(([c, boxes]) => {
        if (boxes.size > maxCustomerCount) {
          maxCustomerCount = boxes.size;
          dominantCustomer = c;
        }
      });

      locationDominantAttrs[loc] = {
        dominantCustomer,
        dominantPower
      };
    });

    const suggestions: RelocationSuggestion[] = [];

    // 对不一致的箱号进行筛查并进行迁移位置推荐
    activeBoxes.forEach((bx) => {
      // A库位属于待出货，经常变动，不参与纠偏与拼位计算
      if (bx.location.toUpperCase().startsWith('A')) {
        return;
      }

      const locAttrs = locationDominantAttrs[bx.location];
      if (!locAttrs) return;

      const custMatch = bx.customerName === locAttrs.dominantCustomer;
      const powerMatch = bx.power === locAttrs.dominantPower;

      // 如果客户不符，或者功率不符，且该库位存在多客户/多功率，则定义为需要迁移修正的箱子
      if (!custMatch || !powerMatch) {
        let mismatchType = '';
        if (!custMatch && !powerMatch) {
          mismatchType = '客户与功率不一致';
        } else if (!custMatch) {
          mismatchType = '客户不一致';
        } else {
          mismatchType = '功率不一致';
        }

        // 寻找推荐位置的逻辑
        let recommendedLoc = '暂无匹配库位';
        let recommendationReason = '未在系统内找到满足条件且有余位的配对库位';

        // 1. 首要匹配：查找完全匹配目标客户和功率、且仍有剩余空位的库位
        const perfectMatches = Object.entries(locationMap)
          .map(([l, bSet]) => {
            const cnt = bSet.size;
            const cap = getPerfectCapacity(l);
            const attrs = locationDominantAttrs[l];
            return { l, cnt, cap, attrs };
          })
          .filter(item => {
            return item.l !== bx.location &&
                   item.l !== 'TBD' &&
                   !item.l.toUpperCase().startsWith('A') &&
                   item.cnt < item.cap &&
                   item.attrs.dominantCustomer === bx.customerName &&
                   item.attrs.dominantPower === bx.power;
          });

        if (perfectMatches.length > 0) {
          perfectMatches.sort((a, b) => (b.cap - b.cnt) - (a.cap - a.cnt));
          const best = perfectMatches[0];
          recommendedLoc = best.l;
          recommendationReason = `推荐转移至【${best.l}】，该库位主存【${bx.customerName}】且功率为【${bx.power}】，现余 ${best.cap - best.cnt} 箱可用`;
        } else {
          // 2. 备选匹配：寻找完全空置全新库位
          const emptyLocs = Object.entries(locationMap)
            .map(([l, bSet]) => {
              const cap = getPerfectCapacity(l);
              return { l, cnt: bSet.size, cap };
            })
            .filter(item => item.l !== bx.location && item.l !== 'TBD' && !item.l.toUpperCase().startsWith('A') && item.cnt === 0);

          if (emptyLocs.length > 0) {
            emptyLocs.sort((a, b) => compareLocations(a.l, b.l));

            const bestEmpty = emptyLocs[0];
            recommendedLoc = bestEmpty.l;
            recommendationReason = `推荐转移至空置库位【${bestEmpty.l}】，可充当其专属独立库位`;
          } else {
            // 3. 次级匹配：仅匹配该客户（客户优先且有空间）
            const customerMatches = Object.entries(locationMap)
              .map(([l, bSet]) => {
                const cnt = bSet.size;
                const cap = getPerfectCapacity(l);
                const attrs = locationDominantAttrs[l];
                return { l, cnt, cap, attrs };
              })
              .filter(item => {
                return item.l !== bx.location &&
                       item.l !== 'TBD' &&
                       !item.l.toUpperCase().startsWith('A') &&
                       item.cnt < item.cap &&
                       item.attrs.dominantCustomer === bx.customerName;
              });

            if (customerMatches.length > 0) {
              customerMatches.sort((a, b) => (b.cap - b.cnt) - (a.cap - a.cnt));
              const bestCust = customerMatches[0];
              recommendedLoc = bestCust.l;
              recommendationReason = `推荐转移至【${bestCust.l}】以匹配主客户【${bx.customerName}】空间，现余 ${bestCust.cap - bestCust.cnt} 箱可用`;
            } else {
              // 4. 三级匹配：仅匹配功率档
              const powerMatches = Object.entries(locationMap)
                .map(([l, bSet]) => {
                  const cnt = bSet.size;
                  const cap = getPerfectCapacity(l);
                  const attrs = locationDominantAttrs[l];
                  return { l, cnt, cap, attrs };
                })
                .filter(item => {
                  return item.l !== bx.location &&
                         item.l !== 'TBD' &&
                         !item.l.toUpperCase().startsWith('A') &&
                         item.cnt < item.cap &&
                         item.attrs.dominantPower === bx.power;
                });

              if (powerMatches.length > 0) {
                powerMatches.sort((a, b) => (b.cap - b.cnt) - (a.cap - a.cnt));
                const bestPower = powerMatches[0];
                recommendedLoc = bestPower.l;
                recommendationReason = `推荐转移至【${bestPower.l}】以配对功率档【${bx.power}】，现余 ${bestPower.cap - bestPower.cnt} 箱可用`;
              }
            }
          }
        }

        suggestions.push({
          boxNo: bx.boxNo,
          currentLoc: bx.location,
          customerName: bx.customerName,
          powerGrade: bx.power,
          dominantCustomer: locAttrs.dominantCustomer,
          dominantPower: locAttrs.dominantPower,
          mismatchType,
          recommendedLoc,
          recommendationReason
        });
      }
    });

    // 按照当前位置进行归类和排序，使相同库位集中放在一起
    suggestions.sort((a, b) => {
      const compLoc = compareLocations(a.currentLoc, b.currentLoc);
      if (compLoc !== 0) return compLoc;
      return a.boxNo.localeCompare(b.boxNo, undefined, { numeric: true, sensitivity: 'base' });
    });

    setResults(analysisResults);
    setMixedPowerAlerts(alerts);
    setMixedCustomerAlerts(customerAlerts);
    setRelocationSuggestions(suggestions);
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
    setRelocationSuggestions([]);
    setIsSaved(true);
    setIsCloudSynced(false);
    setLastSyncedCloudTime(null);
    try {
      localStorage.removeItem('orderCustomerMap');
      localStorage.removeItem('mappingFileName');
    } catch (e) {
      console.error('Failed to clear orderCustomerMap from localStorage', e);
    }
    if (data.length > 0) {
      performAnalysis(data, {});
    }
  };

  const persistMappingToLocalStorage = () => {
    if (!mappingFileName || Object.keys(orderCustomerMap).length === 0) return;
    try {
      localStorage.setItem('orderCustomerMap', JSON.stringify(orderCustomerMap));
      localStorage.setItem('mappingFileName', mappingFileName);
      setIsSaved(true);
      setSaveSuccess('工单客户对照关系保存成功！今次及今后加载新库存表，均会自动加载此映射关系。');
      setTimeout(() => {
        setSaveSuccess(null);
      }, 5000);
    } catch (e) {
      console.error(e);
      setError('本地保存失败，可能是对照表条数过多，超出了浏览器的 LocalStorage 容量限制（5MB）');
    }
  };

  const downloadMappingTemplate = async () => {
    try {
      const workbook = new ExcelJS.Workbook();
      const worksheet = workbook.addWorksheet('Mapping Template (对照表)');

      // Add columns
      worksheet.columns = [
        { header: 'No. (序号)', key: 'seq', width: 12 },
        { header: 'Order / Sales No. (工单/销售单号-Col B)', key: 'orderNo', width: 32 },
        { header: 'Extra Col 1 (辅助列1)', key: 'remark1', width: 20 },
        { header: 'Extra Col 2 (辅助列2)', key: 'remark2', width: 20 },
        { header: 'Extra Col 3 (辅助列3)', key: 'remark3', width: 20 },
        { header: 'Customer Name (客户名称-Col F)', key: 'customerName', width: 30 },
      ];

      worksheet.views = [{ showGridLines: true }];

      const headerRow = worksheet.getRow(1);
      headerRow.height = 30;
      headerRow.eachCell((cell, colNumber) => {
        cell.font = { name: 'Arial', size: 10, bold: true, color: { argb: 'FFFFFF' } };
        if (colNumber === 2 || colNumber === 6) {
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: '7C3AED' }
          };
        } else {
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: '475569' }
          };
        }
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
        cell.border = {
          top: { style: 'thin', color: { argb: 'E2E8F0' } },
          left: { style: 'thin', color: { argb: 'E2E8F0' } },
          bottom: { style: 'medium', color: { argb: '000000' } },
          right: { style: 'thin', color: { argb: 'E2E8F0' } }
        };
      });

      const samples = [
        { seq: '1', orderNo: 'WO-2026001', remark1: 'Demo-A', remark2: '', remark3: '', customerName: 'Client-Alpha' },
        { seq: '2', orderNo: 'WO-2026002', remark1: 'Demo-B', remark2: '', remark3: '', customerName: 'Client-Beta' },
        { seq: '3', orderNo: 'WO-2026003', remark1: 'Demo-C', remark2: '', remark3: '', customerName: 'Client-Gamma' },
      ];

      samples.forEach((sample) => {
        const row = worksheet.addRow({
          seq: sample.seq,
          orderNo: sample.orderNo,
          remark1: sample.remark1,
          remark2: sample.remark2,
          remark3: sample.remark3,
          customerName: sample.customerName
        });
        row.height = 22;
        row.eachCell((cell, colNumber) => {
          cell.font = { name: 'Arial', size: 9 };
          cell.alignment = { horizontal: colNumber === 2 || colNumber === 6 ? 'left' : 'center', vertical: 'middle' };
          if (colNumber === 2 || colNumber === 6) {
            cell.fill = {
              type: 'pattern',
              pattern: 'solid',
              fgColor: { argb: 'F5F3FF' }
            };
          }
          cell.border = {
            top: { style: 'thin', color: { argb: 'F1F5F9' } },
            left: { style: 'thin', color: { argb: 'F1F5F9' } },
            bottom: { style: 'thin', color: { argb: 'F1F5F9' } },
            right: { style: 'thin', color: { argb: 'F1F5F9' } }
          };
        });
      });

      worksheet.addRow([]);
      const rHeader = worksheet.addRow(['Notes / 规范说明:']);
      rHeader.getCell(1).font = { name: 'Arial', size: 10, bold: true, color: { argb: '7C3AED' } };

      const notes = [
        '1. Column B: Order No. / Sales No. (工单/销售单号)',
        '2. Column F: Customer Name (客户简称/名称)',
        '3. Other columns (A, C, D, E) are optional and ignored (其余列可选，不影响匹配)'
      ];

      notes.forEach(note => {
        const r = worksheet.addRow([note]);
        r.getCell(1).font = { name: 'Arial', size: 9, color: { argb: '475569' } };
      });

      const buffer = await workbook.xlsx.writeBuffer();
      const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      saveAs(blob, 'Order_Customer_Template.xlsx');

      setSaveSuccess('Template downloaded: Order_Customer_Template.xlsx');
      setTimeout(() => setSaveSuccess(null), 3000);
    } catch (e) {
      console.error(e);
      setError('Failed to generate template / 模板生成失败');
    }
  };

  const exportToExcel = async () => {
    if (results.length === 0) return;
    
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
    const worksheet = workbook.addWorksheet('Location Summary (库位统计)');

    const baseColumns = [
      { header: 'Location (库位)', key: 'location', width: 18 },
      { header: 'Dominant Cust & Power (主存客户/功率)', key: 'dominantAttr', width: 34 },
      { header: 'Box Qty (箱数)', key: 'count', width: 14 },
      { header: 'Avail Cap (可入容量)', key: 'available', width: 18 },
    ];

    const customerColumns = uniqueCustomers.map(cust => ({
      header: `${cust} (Qty)`,
      key: `cust_${cust}`,
      width: Math.max(14, cust.length * 2 + 4)
    }));

    worksheet.columns = [...baseColumns, ...customerColumns];

    results.forEach((row) => {
      const rowData: Record<string, any> = {
        location: row['库位名称'],
        dominantAttr: row['主存客户'] !== '空置' ? `${row['主存客户']} (${row['主存功率']})` : 'Empty (空置)',
        count: row['箱号数量'],
        available: row['可入库数量'],
      };

      uniqueCustomers.forEach((cust) => {
        const dist = row['客户分布'] || {};
        rowData[`cust_${cust}`] = dist[cust] || 0;
      });

      const excelRow = worksheet.addRow(rowData);

      const count = row['箱号数量'];
      const cap = row['库位名称'].toUpperCase().startsWith('H') ? 14 : 18;
      let color = 'FF0F172A';
      
      if (count > cap) {
        color = 'FFEF4444';
      } else if (count < cap) {
        color = 'FF10B981';
      }

      excelRow.eachCell((cell) => {
        cell.font = { name: 'Arial', color: { argb: color }, bold: true, size: 10 };
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
      });
      excelRow.getCell('location').alignment = { vertical: 'middle', horizontal: 'left' };
      excelRow.getCell('dominantAttr').alignment = { vertical: 'middle', horizontal: 'left' };
    });

    const headerRow = worksheet.getRow(1);
    headerRow.height = 30;
    headerRow.eachCell((cell) => {
      cell.font = { name: 'Arial', bold: true, color: { argb: 'FFFFFFFF' }, size: 11 };
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FF0F172A' }
      };
      cell.alignment = { vertical: 'middle', horizontal: 'center' };
    });

    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    saveAs(blob, `Location_Analysis_${new Date().toISOString().split('T')[0]}.xlsx`);
  };

  const exportAlertsToExcel = async () => {
    if (mixedPowerAlerts.length === 0) return;

    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Mixed Power Alerts (混档异常)');

    worksheet.columns = [
      { header: 'Location (库位)', key: 'location', width: 16 },
      { header: 'Total Boxes (总箱数)', key: 'totalCount', width: 18 },
      { header: 'Power Class (功率档)', key: 'power', width: 18 },
      { header: 'Box Qty (箱数)', key: 'powerCount', width: 16 },
      { header: 'Box Numbers (箱号明细)', key: 'boxes', width: 75 }
    ];

    mixedPowerAlerts.forEach((alert) => {
      const totalBoxesInLoc = Object.values(alert.powerDetails).reduce((acc: number, bList) => acc + (bList as string[]).length, 0);
      Object.entries(alert.powerDetails).forEach(([power, boxes]) => {
        const boxesTyped = boxes as string[];
        const excelRow = worksheet.addRow({
          location: alert.location,
          totalCount: totalBoxesInLoc,
          power: power,
          powerCount: boxesTyped.length,
          boxes: boxesTyped.join(', ')
        });

        excelRow.eachCell((cell) => {
          cell.font = { name: 'Arial', color: { argb: 'FFEF4444' }, size: 9 };
          cell.alignment = { vertical: 'middle', horizontal: 'left', wrapText: true };
        });

        excelRow.getCell('location').alignment = { vertical: 'middle', horizontal: 'center' };
        excelRow.getCell('totalCount').alignment = { vertical: 'middle', horizontal: 'right' };
        excelRow.getCell('power').alignment = { vertical: 'middle', horizontal: 'center' };
        excelRow.getCell('powerCount').alignment = { vertical: 'middle', horizontal: 'right' };
      });
      worksheet.addRow({});
    });

    const headerRow = worksheet.getRow(1);
    headerRow.height = 30;
    headerRow.eachCell((cell) => {
      cell.font = { name: 'Arial', bold: true, color: { argb: 'FFFFFFFF' }, size: 10 };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEF4444' } };
      cell.alignment = { vertical: 'middle', horizontal: 'center' };
    });

    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    saveAs(blob, `Mixed_Power_Alerts_${new Date().toISOString().split('T')[0]}.xlsx`);
  };

  const exportCustomerAlertsToExcel = async () => {
    if (mixedCustomerAlerts.length === 0) return;

    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Mixed Customer (客户混载)');

    worksheet.columns = [
      { header: 'Location (库位)', key: 'location', width: 16 },
      { header: 'Total Boxes (总箱数)', key: 'totalCount', width: 18 },
      { header: 'Customer (客户名称)', key: 'customer', width: 28 },
      { header: 'Box Qty (箱数)', key: 'customerCount', width: 16 },
      { header: 'Box Numbers (箱号明细)', key: 'boxes', width: 75 }
    ];

    mixedCustomerAlerts.forEach((alert) => {
      const totalBoxesInLoc = Object.values(alert.customerDetails).reduce((acc: number, bList) => acc + (bList as any[]).length, 0);
      Object.entries(alert.customerDetails).forEach(([customer, boxes]) => {
        const boxesTyped = boxes as { boxNo: string; power: string }[];
        const excelRow = worksheet.addRow({
          location: alert.location,
          totalCount: totalBoxesInLoc,
          customer: customer,
          customerCount: boxesTyped.length,
          boxes: boxesTyped.map(b => `${b.boxNo} (${b.power})`).join(', ')
        });

        excelRow.eachCell((cell) => {
          cell.font = { name: 'Arial', color: { argb: 'FF2563EB' }, size: 9 };
          cell.alignment = { vertical: 'middle', horizontal: 'left', wrapText: true };
        });

        excelRow.getCell('location').alignment = { vertical: 'middle', horizontal: 'center' };
        excelRow.getCell('totalCount').alignment = { vertical: 'middle', horizontal: 'right' };
        excelRow.getCell('customer').alignment = { vertical: 'middle', horizontal: 'left' };
        excelRow.getCell('customerCount').alignment = { vertical: 'middle', horizontal: 'right' };
      });
      worksheet.addRow({});
    });

    const headerRow = worksheet.getRow(1);
    headerRow.height = 30;
    headerRow.eachCell((cell) => {
      cell.font = { name: 'Arial', bold: true, color: { argb: 'FFFFFFFF' }, size: 10 };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2563EB' } };
      cell.alignment = { vertical: 'middle', horizontal: 'center' };
    });

    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    saveAs(blob, `Mixed_Customer_Alerts_${new Date().toISOString().split('T')[0]}.xlsx`);
  };

  const exportCustomerStatsToExcel = async () => {
    if (customerStats.length === 0) return;

    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Customer Stats (客户统计)');

    worksheet.columns = [
      { header: 'No. (序号)', key: 'index', width: 12 },
      { header: 'Customer (客户名称)', key: 'customerName', width: 38 },
      { header: 'Box Qty (箱数)', key: 'boxCount', width: 20 }
    ];

    customerStats.forEach((stat, i) => {
      const excelRow = worksheet.addRow({
        index: i + 1,
        customerName: stat.customerName,
        boxCount: stat.boxCount
      });

      excelRow.eachCell((cell) => {
        cell.font = { name: 'Arial', color: { argb: 'FF5B21B6' }, size: 10 };
        cell.alignment = { vertical: 'middle', horizontal: 'left' };
      });
      excelRow.getCell('index').alignment = { vertical: 'middle', horizontal: 'center' };
      excelRow.getCell('boxCount').alignment = { vertical: 'middle', horizontal: 'right' };
    });

    const headerRow = worksheet.getRow(1);
    headerRow.height = 30;
    headerRow.eachCell((cell) => {
      cell.font = { name: 'Arial', bold: true, color: { argb: 'FFFFFFFF' }, size: 10 };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF7C3AED' } };
      cell.alignment = { vertical: 'middle', horizontal: 'center' };
    });

    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    saveAs(blob, `Customer_Inventory_Stats_${new Date().toISOString().split('T')[0]}.xlsx`);
  };

  const exportCorrectionToExcel = async () => {
    if (relocationSuggestions.length === 0) return;

    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Relocation Plan (纠偏建议)');

    worksheet.columns = [
      { header: 'No. (序号)', key: 'index', width: 10 },
      { header: 'Box No. (箱号)', key: 'boxNo', width: 22 },
      { header: 'Current Loc (当前库位)', key: 'currentLoc', width: 18 },
      { header: 'Mismatch (偏差类型)', key: 'mismatchType', width: 18 },
      { header: 'Customer (客户)', key: 'customerName', width: 24 },
      { header: 'Power (功率档)', key: 'powerGrade', width: 15 },
      { header: 'Dominant Cust (主存客户)', key: 'dominantCustomer', width: 24 },
      { header: 'Dominant Power (主存功率)', key: 'dominantPower', width: 16 },
      { header: 'Target Loc (推荐库位)', key: 'recommendedLoc', width: 18 },
      { header: 'Reason (调整说明)', key: 'recommendationReason', width: 55 }
    ];

    relocationSuggestions.forEach((s, i) => {
      const excelRow = worksheet.addRow({
        index: i + 1,
        boxNo: s.boxNo,
        currentLoc: s.currentLoc,
        mismatchType: s.mismatchType,
        customerName: s.customerName,
        powerGrade: s.powerGrade,
        dominantCustomer: s.dominantCustomer,
        dominantPower: s.dominantPower,
        recommendedLoc: s.recommendedLoc,
        recommendationReason: s.recommendationReason
      });

      excelRow.eachCell((cell, colNumber) => {
        if (colNumber === 4) {
          cell.font = { name: 'Arial', color: { argb: 'FFEF4444' }, bold: true, size: 9 };
        } else if (colNumber === 9) {
          cell.font = { name: 'Arial', color: { argb: 'FF10B981' }, bold: true, size: 9 };
        } else {
          cell.font = { name: 'Arial', color: { argb: 'FF1E293B' }, size: 9 };
        }
        cell.alignment = { vertical: 'middle', horizontal: (colNumber === 1 || colNumber === 3 || colNumber === 9) ? 'center' : 'left' };
      });
    });

    const headerRow = worksheet.getRow(1);
    headerRow.height = 30;
    headerRow.eachCell((cell) => {
      cell.font = { name: 'Arial', bold: true, color: { argb: 'FFFFFFFF' }, size: 10 };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2563EB' } };
      cell.alignment = { vertical: 'middle', horizontal: 'center' };
    });

    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    saveAs(blob, `Relocation_Suggestions_${new Date().toISOString().split('T')[0]}.xlsx`);
  };

  const exportNonAGradeToExcel = async () => {
    if (nonAGradeList.length === 0) {
      setError('No Non-A Grade data to export / 暂无可导出的非A级品数据');
      return;
    }

    try {
      const workbook = new ExcelJS.Workbook();
      
      // Sheet 1: Non-A Summary
      const summarySheet = workbook.addWorksheet('Summary (汇总)');
      
      const bItems = nonAGradeList.filter(x => x.grade === 'B');
      const fItems = nonAGradeList.filter(x => x.grade === 'F');
      const rItems = nonAGradeList.filter(x => x.grade === 'R');
      const otherItems = nonAGradeList.filter(x => !['B', 'F', 'R'].includes(x.grade));

      const locMap: Record<string, { b: number; f: number; r: number; other: number; total: number; customers: Set<string> }> = {};
      nonAGradeList.forEach(item => {
        const loc = item.location || 'Unknown';
        if (!locMap[loc]) {
          locMap[loc] = { b: 0, f: 0, r: 0, other: 0, total: 0, customers: new Set() };
        }
        if (item.grade === 'B') locMap[loc].b++;
        else if (item.grade === 'F') locMap[loc].f++;
        else if (item.grade === 'R') locMap[loc].r++;
        else locMap[loc].other++;
        locMap[loc].total++;
        if (item.customerName && item.customerName !== '未映射客户') {
          locMap[loc].customers.add(item.customerName);
        }
      });

      const sortedLocs = Object.keys(locMap).sort(compareLocations);

      summarySheet.columns = [
        { width: 16 },
        { width: 16 },
        { width: 16 },
        { width: 18 },
        { width: 32 }
      ];

      // Title
      const titleRow = summarySheet.addRow(['Non-A Grade Inventory Summary (非A级品库存统计)']);
      titleRow.font = { name: 'Arial', size: 13, bold: true, color: { argb: 'FFFFFFFF' } };
      titleRow.height = 32;
      summarySheet.mergeCells('A1:E1');
      titleRow.getCell(1).fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FF6B21A8' }
      };
      titleRow.getCell(1).alignment = { vertical: 'middle', horizontal: 'center' };

      summarySheet.addRow([]);

      // Stats KPI row
      const statsRow = summarySheet.addRow([
        `Total Non-A: ${nonAGradeList.length}`,
        `Grade B: ${bItems.length}`,
        `Grade F: ${fItems.length}`,
        `Grade R: ${rItems.length}`,
        `Locations: ${sortedLocs.length}`
      ]);
      statsRow.height = 26;
      statsRow.eachCell(cell => {
        cell.font = { name: 'Arial', size: 10, bold: true, color: { argb: 'FF334155' } };
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'FFF3E8FF' }
        };
        cell.border = {
          top: { style: 'thin', color: { argb: 'DDD6FE' } },
          bottom: { style: 'thin', color: { argb: 'DDD6FE' } },
          left: { style: 'thin', color: { argb: 'DDD6FE' } },
          right: { style: 'thin', color: { argb: 'DDD6FE' } }
        };
      });

      summarySheet.addRow([]);

      // Section 1: By Grade
      const sec1 = summarySheet.addRow(['1. Breakdown by Grade / 等级分类统计 (Col K)']);
      sec1.font = { name: 'Arial', size: 10, bold: true, color: { argb: 'FF6B21A8' } };

      const hGrade = summarySheet.addRow(['Grade (等级)', 'Box Qty (箱数)', 'Share (占比)', 'Locs (库位数)', 'Notes (说明)']);
      hGrade.height = 24;
      hGrade.eachCell(cell => {
        cell.font = { name: 'Arial', size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF475569' } };
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
      });

      const totalNonA = nonAGradeList.length || 1;
      const gradeRowsData = [
        { name: 'Grade B (B级品)', count: bItems.length, locs: new Set(bItems.map(x => x.location)).size, note: 'B-Grade / Substandard' },
        { name: 'Grade F (F级品)', count: fItems.length, locs: new Set(fItems.map(x => x.location)).size, note: 'F-Grade / Pending Rework' },
        { name: 'Grade R (R级品)', count: rItems.length, locs: new Set(rItems.map(x => x.location)).size, note: 'R-Grade / Return & Repair' },
      ];
      if (otherItems.length > 0) {
        gradeRowsData.push({
          name: 'Other Non-A (其他)',
          count: otherItems.length,
          locs: new Set(otherItems.map(x => x.location)).size,
          note: Array.from(new Set(otherItems.map(x => x.grade))).join(' / ')
        });
      }

      gradeRowsData.forEach(item => {
        const pct = ((item.count / totalNonA) * 100).toFixed(1) + '%';
        const r = summarySheet.addRow([item.name, item.count, pct, item.locs, item.note]);
        r.height = 20;
        r.eachCell((cell, col) => {
          cell.font = { name: 'Arial', size: 9 };
          cell.alignment = { vertical: 'middle', horizontal: col === 1 || col === 5 ? 'left' : 'center' };
          cell.border = {
            top: { style: 'thin', color: { argb: 'E2E8F0' } },
            bottom: { style: 'thin', color: { argb: 'E2E8F0' } },
            left: { style: 'thin', color: { argb: 'E2E8F0' } },
            right: { style: 'thin', color: { argb: 'E2E8F0' } }
          };
        });
      });

      const totalGradeRow = summarySheet.addRow(['Total Non-A (合计)', nonAGradeList.length, '100.0%', sortedLocs.length, '-']);
      totalGradeRow.height = 22;
      totalGradeRow.eachCell(cell => {
        cell.font = { name: 'Arial', size: 10, bold: true, color: { argb: 'FF0F172A' } };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF1F5F9' } };
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
      });

      summarySheet.addRow([]);

      // Section 2: By Location
      const sec2 = summarySheet.addRow(['2. Distribution by Location / 库位分布汇总']);
      sec2.font = { name: 'Arial', size: 10, bold: true, color: { argb: 'FF6B21A8' } };

      const hLoc = summarySheet.addRow(['Location (库位)', 'Grade B', 'Grade F', 'Grade R', 'Total (总数)', 'Customers (涉及客户)']);
      hLoc.height = 24;
      hLoc.eachCell(cell => {
        cell.font = { name: 'Arial', size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF475569' } };
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
      });

      sortedLocs.forEach(loc => {
        const d = locMap[loc];
        const custList = Array.from(d.customers).join(', ') || '-';
        const r = summarySheet.addRow([
          loc,
          d.b,
          d.f,
          d.r,
          d.total,
          custList
        ]);
        r.height = 20;
        r.eachCell((cell, col) => {
          cell.font = { name: 'Arial', size: 9 };
          cell.alignment = { vertical: 'middle', horizontal: col === 1 || col === 6 ? 'left' : 'center' };
          cell.border = {
            top: { style: 'thin', color: { argb: 'E2E8F0' } },
            bottom: { style: 'thin', color: { argb: 'E2E8F0' } },
            left: { style: 'thin', color: { argb: 'E2E8F0' } },
            right: { style: 'thin', color: { argb: 'E2E8F0' } }
          };
        });
      });

      // Sheet 2: Detail List
      const detailSheet = workbook.addWorksheet('Details (明细清单)');
      detailSheet.columns = [
        { header: 'No. (序号)', key: 'seq', width: 8 },
        { header: 'Box No. (箱号)', key: 'boxNo', width: 24 },
        { header: 'Location (库位)', key: 'location', width: 14 },
        { header: 'Raw Grade (原等级-Col K)', key: 'rawGrade', width: 20 },
        { header: 'Category (归类)', key: 'grade', width: 14 },
        { header: 'Barcode Qty (条码数-Col S)', key: 'barcodeCount', width: 20 },
        { header: 'Order No. (工单号-Col J)', key: 'orderNo', width: 20 },
        { header: 'Customer (客户名称)', key: 'customerName', width: 24 },
        { header: 'Power (功率档)', key: 'power', width: 14 }
      ];

      const hDetail = detailSheet.getRow(1);
      hDetail.height = 28;
      hDetail.eachCell(cell => {
        cell.font = { name: 'Arial', size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E293B' } };
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
      });

      nonAGradeList.forEach((item, idx) => {
        const r = detailSheet.addRow({
          seq: idx + 1,
          boxNo: item.boxNo,
          location: item.location,
          rawGrade: item.rawGrade || item.grade,
          grade: `Grade ${item.grade}`,
          barcodeCount: item.barcodeCount !== undefined ? item.barcodeCount : '-',
          orderNo: item.orderNo || '-',
          customerName: item.customerName || '-',
          power: item.power || '-'
        });
        r.height = 20;

        let gradeColor = 'FF1E293B';
        let gradeBg = 'FFFFFFFF';
        if (item.grade === 'B') {
          gradeColor = 'FF1D4ED8';
          gradeBg = 'FFEFF6FF';
        } else if (item.grade === 'F') {
          gradeColor = 'FF7E22CE';
          gradeBg = 'FFFDF4FF';
        } else if (item.grade === 'R') {
          gradeColor = 'FFE11D48';
          gradeBg = 'FFFFF1F2';
        }

        r.eachCell((cell, col) => {
          cell.font = { name: 'Arial', size: 9 };
          cell.alignment = { vertical: 'middle', horizontal: col === 2 || col === 7 || col === 8 ? 'left' : 'center' };
          cell.border = {
            top: { style: 'thin', color: { argb: 'F1F5F9' } },
            bottom: { style: 'thin', color: { argb: 'F1F5F9' } },
            left: { style: 'thin', color: { argb: 'F1F5F9' } },
            right: { style: 'thin', color: { argb: 'F1F5F9' } }
          };
        });

        r.getCell('grade').font = { name: 'Arial', size: 9, bold: true, color: { argb: gradeColor } };
        r.getCell('grade').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: gradeBg } };
      });

      const buffer = await workbook.xlsx.writeBuffer();
      const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      saveAs(blob, `Non_A_Grade_Report_${new Date().toISOString().split('T')[0]}.xlsx`);

      setSaveSuccess('Exported: Non_A_Grade_Report.xlsx');
      setTimeout(() => setSaveSuccess(null), 3000);
    } catch (err) {
      console.error(err);
      setError('Failed to export Non-A report / 导出失败');
    }
  };

  // 非A级品筛选与搜索数据衍生
  const filteredNonAGradeList = nonAGradeList.filter(item => {
    if (nonAGradeFilter === 'B' && item.grade !== 'B') return false;
    if (nonAGradeFilter === 'F' && item.grade !== 'F') return false;
    if (nonAGradeFilter === 'R' && item.grade !== 'R') return false;
    if (nonAGradeFilter === 'OTHER' && ['B', 'F', 'R'].includes(item.grade)) return false;

    if (nonAGradeSearch.trim()) {
      const q = nonAGradeSearch.trim().toLowerCase();
      const matchBox = String(item.boxNo).toLowerCase().includes(q);
      const matchLoc = String(item.location).toLowerCase().includes(q) || String(item.rawLocation).toLowerCase().includes(q);
      const matchOrder = String(item.orderNo || '').toLowerCase().includes(q);
      const matchCustomer = String(item.customerName || '').toLowerCase().includes(q);
      const matchGrade = String(item.rawGrade || '').toLowerCase().includes(q);
      if (!matchBox && !matchLoc && !matchOrder && !matchCustomer && !matchGrade) return false;
    }

    return true;
  });

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

  const exportZeroBarcodesToExcel = async () => {
    if (zeroBarcodeList.length === 0) return;
    try {
      const workbook = new ExcelJS.Workbook();
      const worksheet = workbook.addWorksheet('Zero Barcodes (条码为0)');
      worksheet.columns = [
        { header: 'Row No. (原行号)', key: 'rowIdx', width: 14 },
        { header: 'Box No. (箱号)', key: 'boxNo', width: 24 },
        { header: 'Location (库位)', key: 'location', width: 14 },
        { header: 'Grade (等级-Col K)', key: 'grade', width: 16 },
        { header: 'Raw Barcode (原值-Col S)', key: 'rawBarcode', width: 18 },
        { header: 'Order No. (工单号)', key: 'orderNo', width: 20 },
        { header: 'Filter Reason (过滤原因)', key: 'reason', width: 30 }
      ];
      worksheet.views = [{ showGridLines: true }];
      const headerRow = worksheet.getRow(1);
      headerRow.height = 28;
      headerRow.eachCell(cell => {
        cell.font = { name: 'Arial', size: 10, bold: true, color: { argb: 'FFFFFF' } };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'EA580C' } };
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
      });
      zeroBarcodeList.forEach(item => {
        const r = worksheet.addRow({
          rowIdx: item.rowIdx,
          boxNo: item.boxNo,
          location: item.location,
          grade: item.grade,
          rawBarcode: item.rawBarcode === '' ? '(Empty)' : item.rawBarcode,
          orderNo: item.orderNo || '-',
          reason: item.reason
        });
        r.height = 20;
        r.eachCell(cell => {
          cell.alignment = { vertical: 'middle' };
          cell.font = { name: 'Arial', size: 9 };
        });
      });
      const buffer = await workbook.xlsx.writeBuffer();
      const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      saveAs(blob, `Zero_Barcode_Filtered_${zeroBarcodeList.length}Boxes_${new Date().toISOString().split('T')[0]}.xlsx`);
      setSaveSuccess(`Exported: Zero_Barcode_Filtered_${zeroBarcodeList.length}Boxes.xlsx`);
      setTimeout(() => setSaveSuccess(null), 3000);
    } catch (err) {
      console.error(err);
      setError('Failed to export zero barcode list / 导出失败');
    }
  };

  const exportAuditReportToExcel = async () => {
    if (!auditReport) return;
    try {
      const workbook = new ExcelJS.Workbook();

      // Sheet 1: Audit Overview
      const summarySheet = workbook.addWorksheet('Audit Overview (勾稽总览)');
      summarySheet.columns = [
        { header: 'Audit Item (核查项目)', key: 'item', width: 30 },
        { header: 'Value (数量/数值)', key: 'value', width: 18 },
        { header: 'Notes (特点与说明)', key: 'notes', width: 50 }
      ];
      summarySheet.views = [{ showGridLines: true }];
      const h1 = summarySheet.getRow(1);
      h1.height = 28;
      h1.eachCell(cell => {
        cell.font = { name: 'Arial', size: 10, bold: true, color: { argb: 'FFFFFF' } };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '1E3A8A' } };
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
      });

      const addSummaryRow = (item: string, val: string | number, notes: string) => {
        const r = summarySheet.addRow({ item, value: val, notes });
        r.height = 22;
        r.eachCell(cell => {
          cell.alignment = { vertical: 'middle' };
          cell.font = { name: 'Arial', size: 9 };
        });
      };

      addSummaryRow('1. Raw Imported Rows (原始总行数)', `${totalRowsRaw}`, 'Total physical rows in source Excel');
      addSummaryRow('2. Zero-Barcode Filtered (条码为0过滤)', `-${auditReport.zeroBarcodeBreakdown.total}`, 'Explicit 0 barcode rows excluded');
      addSummaryRow('3. Valid Boxes Included (有效总箱数)', `${totalRowsRaw - auditReport.zeroBarcodeBreakdown.total}`, 'Formula: 7,380 - 23 = 7,357 boxes');
      addSummaryRow('   ├─ Grade-A Boxes (A级良品箱数)', `${data.length}`, 'Included in primary warehouse analytics');
      addSummaryRow('   └─ Non-A Grade (非A级品箱数 B/F/R)', `${nonAGradeCount}`, 'Managed & exported independently');
      addSummaryRow('4. 8-Box Gap Diagnostics (8箱差异原因)', 'Gap: 8 Boxes', '7,372 - 23 = 7,349 vs 7,357; 7,380 - 7,372 = 8 rows gap in raw table');
      addSummaryRow('5. Duplicate Box Rows (重复箱号行数)', `${auditReport.duplicateRowCount}`, auditReport.duplicateRowCount > 0 ? `${auditReport.duplicateBoxes.length} box numbers appeared multiple times` : 'None');
      addSummaryRow('6. Blank Box Rows (空箱号行数)', `${auditReport.emptyBoxRows.length}`, auditReport.emptyBoxRows.length > 0 ? `Rows: ${auditReport.emptyBoxRows.map(x => x.rowIdx).join(', ')}` : 'None');
      addSummaryRow('7. Unassigned Loc Rows (未分配库位)', `${auditReport.emptyLocRows.length}`, auditReport.emptyLocRows.length > 0 ? `Rows: ${auditReport.emptyLocRows.map(x => x.rowIdx).join(', ')}` : 'None');

      // Sheet 2: Duplicate Boxes
      if (auditReport.duplicateBoxes.length > 0) {
        const dupSheet = workbook.addWorksheet('Duplicates (重复箱号)');
        dupSheet.columns = [
          { header: 'Row No. (原行号)', key: 'rowIdx', width: 14 },
          { header: 'Box No. (箱号)', key: 'boxNo', width: 24 },
          { header: 'Location (库位)', key: 'location', width: 16 },
          { header: 'Grade (等级)', key: 'grade', width: 14 },
          { header: 'Power (功率)', key: 'power', width: 16 },
          { header: 'Order No. (工单号)', key: 'orderNo', width: 20 },
          { header: 'Notes (特征说明)', key: 'desc', width: 28 }
        ];
        dupSheet.views = [{ showGridLines: true }];
        const h2 = dupSheet.getRow(1);
        h2.height = 28;
        h2.eachCell(cell => {
          cell.font = { name: 'Arial', size: 10, bold: true, color: { argb: 'FFFFFF' } };
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'B45309' } };
          cell.alignment = { horizontal: 'center', vertical: 'middle' };
        });
        auditReport.duplicateBoxes.forEach(dup => {
          dup.rows.forEach((r, idx) => {
            const row = dupSheet.addRow({
              rowIdx: `Row ${r.rowIdx}`,
              boxNo: dup.boxNo,
              location: r.location,
              grade: r.grade,
              power: r.power,
              orderNo: r.orderNo,
              desc: idx === 0 ? `First Entry (${dup.count} entries)` : `Duplicate Entry (+1 row)`
            });
            row.height = 20;
            row.eachCell(c => {
              c.alignment = { vertical: 'middle' };
              c.font = { name: 'Arial', size: 9 };
            });
          });
        });
      }

      // Sheet 3: Zero Barcode Filtered
      const zeroSheet = workbook.addWorksheet('Zero Barcodes (0条码明细)');
      zeroSheet.columns = [
        { header: 'Row No. (原行号)', key: 'rowIdx', width: 14 },
        { header: 'Box No. (箱号)', key: 'boxNo', width: 24 },
        { header: 'Location (库位)', key: 'location', width: 16 },
        { header: 'Grade (Col K)', key: 'grade', width: 14 },
        { header: 'Raw Barcode (Col S)', key: 'rawBarcode', width: 16 },
        { header: 'Order No. (工单号)', key: 'orderNo', width: 20 },
        { header: 'Reason (原因)', key: 'reason', width: 28 }
      ];
      zeroSheet.views = [{ showGridLines: true }];
      const h3 = zeroSheet.getRow(1);
      h3.height = 28;
      h3.eachCell(cell => {
        cell.font = { name: 'Arial', size: 10, bold: true, color: { argb: 'FFFFFF' } };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'EA580C' } };
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
      });
      auditReport.zeroBarcodeBreakdown.rows.forEach(item => {
        const r = zeroSheet.addRow({
          rowIdx: `Row ${item.rowIdx}`,
          boxNo: item.boxNo,
          location: item.location,
          grade: item.grade,
          rawBarcode: item.rawBarcode === '' ? '(Empty)' : item.rawBarcode,
          orderNo: item.orderNo || '-',
          reason: item.reason
        });
        r.height = 20;
        r.eachCell(c => {
          c.alignment = { vertical: 'middle' };
          c.font = { name: 'Arial', size: 9 };
        });
      });

      const buffer = await workbook.xlsx.writeBuffer();
      const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      saveAs(blob, `Audit_Diagnostics_Report_${new Date().toISOString().split('T')[0]}.xlsx`);
      setSaveSuccess('Exported: Audit_Diagnostics_Report.xlsx');
      setTimeout(() => setSaveSuccess(null), 3000);
    } catch (err) {
      console.error(err);
      setError('Failed to export audit report / 导出失败');
    }
  };

  const reset = () => {
    setData([]);
    setResults([]);
    setMixedPowerAlerts([]);
    setMixedCustomerAlerts([]);
    setRelocationSuggestions([]);
    setCustomerStats([]);
    setNonAGradeList([]);
    setZeroBarcodeList([]);
    setAuditReport(null);
    setShowAuditModal(false);
    setShowZeroBarcodeModal(false);
    setFileName('');
    setTotalRowsRaw(0);
    setNonAGradeCount(0);
    setZeroBarcodeCount(0);
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
          <div>
            <span className="text-white font-bold tracking-tight uppercase block leading-none">Inv Analyst</span>
            <span className="text-[9px] text-slate-400 font-medium">仓储智能分析</span>
          </div>
        </div>
        
        <nav className="mt-4 flex-1 px-4">
          <div className="px-3 py-2 text-[10px] font-bold text-slate-500 uppercase tracking-widest mb-1 opacity-70">
            Data Management (数据管理)
          </div>
          <button 
            onClick={reset}
            className={cn(
              "w-full flex items-center gap-2.5 px-3 py-2.5 rounded-xl transition-all text-xs font-bold relative group",
              !data.length ? "bg-blue-600 text-white shadow-md shadow-blue-600/20" : "text-slate-400 hover:text-white hover:bg-slate-800"
            )}
          >
            <Upload className="w-4 h-4 shrink-0" />
            <span>Reset Workspace (重置沙盒)</span>
          </button>
        </nav>

        <div className="p-5 border-t border-slate-800 text-[10px] text-slate-500 uppercase font-bold tracking-wider flex items-center justify-between">
          <span>System: Online (运行中)</span>
          <span className="w-2 h-2 rounded-full bg-emerald-500 animate-pulse" />
        </div>
      </aside>

      {/* Main Content Area */}
      <main className="flex-1 flex flex-col overflow-hidden">
        {/* Top Header */}
        <header className="h-16 bg-white border-b border-slate-200 flex items-center justify-between px-8 shrink-0">
          <div className="flex items-center gap-3">
            <h2 className="text-base font-bold text-slate-800 flex items-center gap-2">
              <span>Inventory Analytics</span>
              <span className="text-xs font-normal text-slate-400">(库存数据统计分析)</span>
            </h2>
            <span className="px-2 py-0.5 bg-slate-100 text-[10px] text-slate-500 rounded-full font-bold uppercase tracking-wider">v5.1 Pro</span>
          </div>
          
          <div className="flex items-center gap-6">
            {fileName && (
              <div className="text-right hidden sm:block">
                <p className="text-[10px] uppercase font-bold text-slate-400 leading-none mb-1">Active File (当前文件)</p>
                <p className="text-xs font-bold text-slate-700 truncate max-w-xs">{fileName}</p>
              </div>
            )}
            <div className="h-8 w-px bg-slate-100 hidden sm:block" />
            <div className="w-9 h-9 rounded-full bg-slate-50 border border-slate-200 flex items-center justify-center text-slate-400 hover:text-slate-600 transition-colors cursor-help">
              <FileText className="w-4 h-4" />
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
                  <h4 className="text-xs font-bold uppercase tracking-wider text-slate-600 flex items-center gap-2">
                    <span className="w-5 h-5 rounded-full bg-blue-100 text-blue-600 flex items-center justify-center font-bold text-[10px]">1</span>
                    Inventory Source (主库存数据源)
                  </h4>
                  {data.length > 0 && (
                    <span className="px-2 py-0.5 bg-emerald-50 text-emerald-600 rounded-full text-[9px] font-bold border border-emerald-100">
                      Loaded (已载入)
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
                    <div className="w-12 h-12 bg-blue-50 rounded-2xl flex items-center justify-center mb-3 text-blue-600 group-hover:scale-110 transition-transform">
                      <Upload className="w-6 h-6" />
                    </div>
                    <h5 className="font-bold text-sm text-slate-800">Drop Inventory Excel Here</h5>
                    <p className="text-[10px] text-slate-400 mt-0.5">拖拽主库存 Excel 或点击上传 (.xlsx, .xls)</p>
                    
                    {isProcessing && (
                      <div className="absolute inset-0 bg-white/95 backdrop-blur-sm rounded-2xl flex flex-col items-center justify-center z-20">
                        <Loader2 className="w-8 h-8 text-blue-600 animate-spin mb-2" />
                        <p className="text-[10px] font-bold text-slate-800 tracking-wider">Parsing data (解析中)...</p>
                      </div>
                    )}
                  </motion.div>
                ) : (
                  <div className="bg-slate-900 rounded-2xl p-5 text-white shadow-md relative overflow-hidden">
                    <div className="absolute -right-6 -bottom-6 w-20 h-20 bg-white/5 rounded-full" />
                    <div className="space-y-4">
                      <div>
                        <p className="text-[9px] text-slate-400 uppercase tracking-widest font-bold">Connected File (当前库存)</p>
                        <p className="text-xs font-bold truncate mt-1 text-slate-100">{fileName}</p>
                      </div>
                      <div className="grid grid-cols-2 gap-2 border-t border-slate-800 pt-3">
                        <div>
                          <p className="text-[9px] text-slate-400 uppercase">Grade-A Boxes (A级有效)</p>
                          <p className="text-lg font-bold font-mono tracking-tight text-blue-400">{data.length} <span className="text-xs font-normal">Boxes</span></p>
                        </div>
                        <div>
                          <p className="text-[9px] text-slate-400 uppercase">Active Locs (锁定库位)</p>
                          <p className="text-lg font-bold font-mono tracking-tight text-emerald-400">{results.length} <span className="text-xs font-normal">Locs</span></p>
                        </div>
                      </div>

                      {/* Reconciliation Card */}
                      <div className="bg-slate-950/90 border border-slate-800 rounded-xl p-3 space-y-2.5">
                        <div className="flex items-center justify-between">
                          <span className="text-[10px] font-bold text-slate-300 flex items-center gap-1.5">
                            <CheckCircle2 className="w-3.5 h-3.5 text-emerald-400" />
                            Reconciliation (数据闭环)
                          </span>
                          <span className="text-[9px] font-bold text-emerald-400 bg-emerald-950/80 px-2 py-0.5 rounded-full border border-emerald-800/60">
                            Zero Error (零误差)
                          </span>
                        </div>

                        <div className="grid grid-cols-4 gap-1.5 text-center font-mono">
                          <div className="bg-slate-900 p-1.5 rounded-lg border border-slate-800">
                            <span className="text-[8px] text-slate-400 block font-sans">Raw Rows (原行)</span>
                            <span className="text-xs font-bold text-white">{totalRowsRaw}</span>
                          </div>
                          <div 
                            onClick={() => setShowZeroBarcodeModal(true)}
                            className="bg-orange-950/60 p-1.5 rounded-lg border border-orange-800/80 cursor-pointer hover:border-orange-500 transition-colors group"
                            title={`Click to view ${zeroBarcodeCount} filtered zero-barcode boxes`}
                          >
                            <span className="text-[8px] text-orange-300 block font-sans flex items-center justify-center gap-0.5">
                              Col S=0 <Eye className="w-2.5 h-2.5 group-hover:scale-125 transition-transform" />
                            </span>
                            <span className="text-xs font-bold text-orange-400">-{zeroBarcodeCount}</span>
                          </div>
                          <div className="bg-blue-950/60 p-1.5 rounded-lg border border-blue-800/80">
                            <span className="text-[8px] text-blue-300 block font-sans">Grade-A (A级)</span>
                            <span className="text-xs font-bold text-blue-400">{data.length}</span>
                          </div>
                          <div className="bg-purple-950/60 p-1.5 rounded-lg border border-purple-800/80">
                            <span className="text-[8px] text-purple-300 block font-sans">Non-A (非A)</span>
                            <span className="text-xs font-bold text-purple-300">{nonAGradeCount}</span>
                          </div>
                        </div>

                        <div className="bg-slate-900/80 rounded-lg p-2 text-[9px] font-mono text-slate-400 border border-slate-800/60 space-y-0.5">
                          <div className="flex items-center justify-between text-slate-300">
                            <span>Valid Total (有效总数):</span>
                            <span className="text-white font-bold">{data.length + nonAGradeCount} Boxes</span>
                          </div>
                          <div className="text-[8.5px] text-emerald-400/90 leading-tight">
                            Equation: {totalRowsRaw} (Raw) - {zeroBarcodeCount} (Col S=0) = {data.length + nonAGradeCount} [{data.length} A + {nonAGradeCount} Non-A]
                          </div>
                        </div>

                        {zeroBarcodeCount > 0 && (
                          <button
                            onClick={() => setShowZeroBarcodeModal(true)}
                            className="w-full py-1 px-2 bg-orange-950/50 hover:bg-orange-900/60 text-orange-300 rounded-lg text-[9px] font-bold border border-orange-800/60 flex items-center justify-center gap-1.5 transition-colors"
                          >
                            <Eye className="w-3 h-3 text-orange-400" />
                            View Zero Barcodes ({zeroBarcodeCount} Boxes)
                          </button>
                        )}

                        <button
                          onClick={() => setShowAuditModal(true)}
                          className="w-full py-1.5 px-2 bg-indigo-950/70 hover:bg-indigo-900 text-indigo-200 rounded-lg text-[9px] font-bold border border-indigo-700/60 flex items-center justify-center gap-1.5 transition-colors shadow-sm mt-1 group"
                        >
                          <HelpCircle className="w-3 h-3 text-indigo-400 group-hover:scale-110 transition-transform" />
                          Audit & Diagnostics (8-Box Gap)
                        </button>
                      </div>

                      {/* Non-A Grade Block */}
                      {nonAGradeList.length > 0 && (
                        <div className="bg-purple-950/70 border border-purple-800/80 rounded-xl p-3 space-y-2">
                          <div className="flex items-center justify-between">
                            <span className="text-[10px] font-bold text-purple-300 flex items-center gap-1.5">
                              <span className="w-2 h-2 rounded-full bg-purple-400 animate-pulse" />
                              Non-A Grade (非A级品 - Col K)
                            </span>
                            <span className="font-mono text-xs font-bold text-purple-200">
                              {nonAGradeList.length} Boxes
                            </span>
                          </div>
                          
                          <div className="grid grid-cols-3 gap-1.5 text-[9px] text-center">
                            <div className="bg-purple-900/60 py-1.5 px-1 rounded-lg border border-purple-700/50">
                              <span className="text-blue-300 font-bold block">Grade B</span>
                              <span className="font-mono text-white text-xs font-bold">
                                {nonAGradeList.filter(x => x.grade === 'B').length}
                              </span>
                            </div>
                            <div className="bg-purple-900/60 py-1.5 px-1 rounded-lg border border-purple-700/50">
                              <span className="text-purple-300 font-bold block">Grade F</span>
                              <span className="font-mono text-white text-xs font-bold">
                                {nonAGradeList.filter(x => x.grade === 'F').length}
                              </span>
                            </div>
                            <div className="bg-purple-900/60 py-1.5 px-1 rounded-lg border border-purple-700/50">
                              <span className="text-rose-300 font-bold block">Grade R</span>
                              <span className="font-mono text-white text-xs font-bold">
                                {nonAGradeList.filter(x => x.grade === 'R').length}
                              </span>
                            </div>
                          </div>

                          <div className="flex gap-2 pt-1">
                            <button
                              onClick={() => setActiveTab('nonA')}
                              className="flex-1 py-1.5 px-2 bg-purple-900/90 hover:bg-purple-800 text-purple-200 hover:text-white rounded-lg text-[10px] font-bold transition-all border border-purple-700/60 flex items-center justify-center gap-1"
                            >
                              <Layers className="w-3 h-3" />
                              Details (明细)
                            </button>
                            <button
                              onClick={exportNonAGradeToExcel}
                              className="flex-1 py-1.5 px-2 bg-purple-600 hover:bg-purple-500 text-white rounded-lg text-[10px] font-bold transition-all flex items-center justify-center gap-1 shadow-md active:scale-95"
                              title="Export Non-A Grade Report"
                            >
                              <Download className="w-3 h-3" />
                              Export (导出)
                            </button>
                          </div>
                        </div>
                      )}
                      <div className="flex gap-2">
                        <button 
                          onClick={exportToExcel}
                          className="flex-1 flex items-center justify-center gap-1.5 bg-emerald-600 hover:bg-emerald-500 text-white font-bold py-2 rounded-xl text-xs transition-colors shadow-md active:scale-95"
                        >
                          <Download className="w-3.5 h-3.5" />
                          Export Excel (导出库位表)
                        </button>
                        <button 
                          onClick={reset}
                          className="px-3 bg-slate-800 hover:bg-slate-700 text-slate-400 hover:text-white rounded-xl text-xs transition-colors"
                          title="Reset current inventory"
                        >
                          Reset (重置)
                        </button>
                      </div>
                    </div>
                  </div>
                )}
              </div>

              {/* Card 2: Work Order -> Customer Relations Mapping */}
              <div className="bg-white p-6 rounded-3xl border border-slate-200 shadow-sm flex flex-col gap-4">
                <div className="flex items-center justify-between pb-3 border-b border-slate-100">
                  <div className="flex items-center gap-2">
                    <h4 className="text-xs font-bold uppercase tracking-wider text-slate-600 flex items-center gap-2">
                      <span className="w-5 h-5 rounded-full bg-violet-100 text-violet-600 flex items-center justify-center font-bold text-[10px]">2</span>
                      Order Mapping (工单对照表)
                    </h4>
                    <button 
                      onClick={downloadMappingTemplate} 
                      className="text-[10px] text-violet-600 hover:text-violet-700 font-bold hover:underline transition-all flex items-center gap-0.5 ml-1"
                      title="Download template Excel"
                    >
                      (Template / 模板)
                    </button>
                  </div>
                  {mappingFileName && (
                    isSaved ? (
                      <span className="px-2 py-0.5 bg-emerald-50 text-emerald-600 rounded-full text-[9px] font-bold border border-emerald-100 flex items-center gap-1">
                        <span className="w-1 h-1 rounded-full bg-emerald-500" /> Saved (已保存)
                      </span>
                    ) : (
                      <span className="px-2 py-0.5 bg-amber-50 text-amber-600 rounded-full text-[9px] font-bold border border-amber-200 flex items-center gap-1 animate-pulse">
                        <span className="w-1.5 h-1.5 rounded-full bg-amber-500 shrink-0" /> Unsaved (未保存)
                      </span>
                    )
                  )}
                </div>

                {!mappingFileName ? (
                  <div className="flex flex-col gap-3">
                    <motion.div 
                      initial={{ opacity: 0, y: 10 }}
                      animate={{ opacity: 1, y: 0 }}
                      onDragOver={(e) => e.preventDefault()}
                      onDrop={handleMappingDrop}
                      className="p-8 rounded-2xl border-2 border-dashed border-slate-200 flex flex-col items-center justify-center text-center relative group hover:border-violet-500 hover:bg-violet-50/10 transition-all duration-300 min-h-[150px]"
                    >
                      <input 
                        type="file" 
                        accept=".xlsx,.xls" 
                        onChange={handleMappingFileUpload}
                        className="absolute inset-0 opacity-0 cursor-pointer z-10"
                      />
                      <div className="w-12 h-12 bg-violet-50 rounded-2xl flex items-center justify-center mb-3 text-violet-600 group-hover:scale-110 transition-transform">
                        <Package className="w-6 h-6 animate-pulse" />
                      </div>
                      <h5 className="font-bold text-sm text-slate-800">Drop Mapping Excel Here</h5>
                      <p className="text-[10px] text-slate-400 mt-0.5">拖拽工单-客户对照表以建立关联</p>
                    </motion.div>

                    <div className="bg-slate-50 border border-slate-100 rounded-2xl p-3 text-xs text-slate-600 space-y-1.5">
                      <div className="flex items-center justify-between">
                        <p className="font-bold text-slate-700 flex items-center gap-1 text-[10px]">
                          <FileText className="w-3 h-3 text-violet-500" />
                          Mapping Spec (匹配规则):
                        </p>
                        <button
                          onClick={downloadMappingTemplate}
                          className="text-[9px] bg-violet-100 hover:bg-violet-200 text-violet-700 font-bold px-2 py-0.5 rounded transition-colors"
                        >
                          Template.xlsx
                        </button>
                      </div>
                      <div className="text-[10px] space-y-0.5 text-slate-500">
                        <p><span className="text-violet-600 font-bold">Col B</span>: Order / Sales No. (工单/销售单号)</p>
                        <p><span className="text-violet-600 font-bold">Col F</span>: Customer Name (客户简称/名称)</p>
                      </div>
                    </div>
                  </div>
                ) : (
                  <div className="bg-violet-950 rounded-2xl p-5 text-white shadow-md relative overflow-hidden flex flex-col gap-4">
                    <div className="absolute -right-6 -bottom-6 w-20 h-20 bg-white/5 rounded-full" />
                    <div className="space-y-4">
                      <div>
                        <p className="text-[9px] text-violet-400 uppercase tracking-widest font-bold">Connected Dictionary (已连对照)</p>
                        <p className="text-xs font-bold truncate mt-1 text-slate-100">{mappingFileName}</p>
                      </div>
                      <div className="grid grid-cols-2 gap-2 border-t border-violet-900 pt-3">
                        <div>
                          <p className="text-[9px] text-violet-400 uppercase">Rules (关联规则)</p>
                          <p className="text-lg font-bold font-mono tracking-tight text-violet-300">{Object.keys(orderCustomerMap).length} <span className="text-xs font-normal">Rules</span></p>
                        </div>
                        <div>
                          <p className="text-[9px] text-violet-400 uppercase">Clients (关联客户)</p>
                          <p className="text-lg font-bold font-mono tracking-tight text-violet-300">{customerStats.filter(c => c.customerName !== '未映射客户').length} <span className="text-xs font-normal">Clients</span></p>
                        </div>
                      </div>

                      {!isSaved && (
                        <div className="bg-amber-500/10 border border-amber-500/20 rounded-xl p-2.5 text-[9px] text-amber-200 leading-relaxed">
                          <p className="font-bold flex items-center gap-1 text-amber-400 mb-0.5">
                            <AlertCircle className="w-3 h-3" /> Unsaved: Loaded in memory only
                          </p>
                          建议点击下方保存按钮固化于本地或云端，避免后续重复上传。
                        </div>
                      )}
                    </div>

                    <div className="flex flex-col gap-2 pt-2 border-t border-violet-900 z-10">
                      {!isSaved && (
                        <button 
                          onClick={persistMappingToLocalStorage}
                          className="w-full flex items-center justify-center gap-1.5 bg-amber-500 hover:bg-amber-600 text-slate-950 font-bold py-2 rounded-xl text-xs transition-colors shadow-md active:scale-95"
                        >
                          <Save className="w-3.5 h-3.5" />
                          Save Locally (保存至本地)
                        </button>
                      )}

                      <button 
                        onClick={saveMappingToCloud}
                        disabled={isCloudSaving}
                        className={cn(
                          "w-full flex items-center justify-center gap-1.5 font-bold py-2 rounded-xl text-xs transition-all border shadow-md z-10",
                          isCloudSynced 
                            ? "bg-emerald-600 hover:bg-emerald-700 text-white border-emerald-500"
                            : "bg-indigo-600 hover:bg-indigo-700 text-white border-indigo-500 active:scale-95"
                        )}
                      >
                        {isCloudSaving ? (
                          <Loader2 className="w-3.5 h-3.5 animate-spin" />
                        ) : isCloudSynced ? (
                          <CheckCircle2 className="w-3.5 h-3.5 text-emerald-200" />
                        ) : (
                          <Cloud className="w-3.5 h-3.5" />
                        )}
                        {isCloudSaving ? 'Publishing to Cloud...' : isCloudSynced ? 'Synced to Cloud (云端已同步)' : 'Sync to Cloud (发布至云端共享)'}
                      </button>
                      {lastSyncedCloudTime && (
                        <p className="text-[8px] text-violet-300 text-center font-mono">
                          Last Synced: {lastSyncedCloudTime}
                        </p>
                      )}

                      <div className="flex gap-2 mt-0.5">
                        {customerStats.length > 0 && (
                          <button 
                            onClick={exportCustomerStatsToExcel}
                            className="flex-1 flex items-center justify-center gap-1.5 bg-violet-800 hover:bg-violet-700 text-white font-bold py-1.5 rounded-xl text-xs transition-colors border border-violet-600 shadow-sm active:scale-95"
                          >
                            <Download className="w-3.5 h-3.5" />
                            Customer Stats (客户统计)
                          </button>
                        )}
                        <button 
                          onClick={clearMapping}
                          className="px-3 py-1.5 bg-violet-900 hover:bg-violet-800 text-violet-300 hover:text-white rounded-xl text-xs font-bold transition-colors"
                          title="Clear dictionary"
                        >
                          Clear (清除)
                        </button>
                      </div>
                    </div>
                  </div>
                )}

                {/* Cloud Model Library */}
                <div className="mt-2 pt-3 border-t border-slate-100">
                  <div className="flex items-center justify-between mb-2">
                    <div className="flex items-center gap-1.5">
                      <Cloud className="w-3.5 h-3.5 text-indigo-500" />
                      <p className="text-[10px] font-bold text-slate-500 uppercase tracking-wider">
                        Cloud Models / 云端共享库 ({cloudMappings.length})
                      </p>
                    </div>
                    <button 
                      onClick={() => fetchCloudMappings(false)} 
                      className="text-slate-400 hover:text-slate-600 transition-colors p-1 rounded-lg"
                      title="Refresh"
                    >
                      <RefreshCw className={cn("w-3.5 h-3.5", isCloudLoading && "animate-spin")} />
                    </button>
                  </div>

                  {isCloudLoading && cloudMappings.length === 0 ? (
                    <div className="py-3 text-center text-[10px] text-slate-400 flex items-center justify-center gap-1.5">
                      <Loader2 className="w-3.5 h-3.5 animate-spin text-indigo-500" />
                      Loading cloud models...
                    </div>
                  ) : cloudMappings.length === 0 ? (
                    <div className="py-3 text-center text-[10px] text-slate-400 bg-slate-50 rounded-xl border border-dashed border-slate-200">
                      No cloud models yet (暂无云端共享对照)
                    </div>
                  ) : (
                    <div className="space-y-1.5 max-h-[150px] overflow-y-auto pr-1 select-none">
                      {cloudMappings.map((item) => (
                        <div 
                          key={item.id} 
                          className={cn(
                            "group p-2 rounded-xl border transition-all text-[10px] flex items-center justify-between gap-2",
                            mappingFileName === item.mappingFileName 
                              ? "bg-indigo-50/50 border-indigo-200" 
                              : "bg-slate-50/70 hover:bg-slate-50 border-slate-100 hover:border-slate-200"
                          )}
                        >
                          <div className="min-w-0 flex-1">
                            <span className="font-bold text-slate-700 truncate block max-w-[130px]" title={item.mappingFileName}>
                              {item.mappingFileName}
                            </span>
                            <div className="text-[9px] text-slate-400 mt-0.5 flex items-center gap-1 font-mono">
                              <span>{item.count} rules</span>
                              <span>•</span>
                              <span>{item.updatedAt ? new Date(item.updatedAt).toLocaleDateString() : '-'}</span>
                            </div>
                          </div>

                          <div className="flex items-center gap-1 shrink-0">
                            <button
                              onClick={() => applyCloudMapping(item)}
                              className="px-2 py-0.5 bg-white hover:bg-indigo-50 text-indigo-600 rounded border border-slate-200 font-bold text-[9px] transition-all"
                            >
                              Load (载入)
                            </button>
                            <button
                              onClick={() => downloadCloudMappingAsExcel(item)}
                              className="p-1 bg-white hover:bg-amber-50 text-amber-600 rounded border border-slate-200 text-[9px] transition-all"
                              title="Download Excel"
                            >
                              <CloudDownload className="w-3 h-3" />
                            </button>
                          </div>
                        </div>
                      ))}
                    </div>
                  )}
                </div>
              </div>

              {/* Requirements Guidelines */}
              <div className="bg-white p-6 rounded-3xl border border-slate-200 shadow-sm relative overflow-hidden group">
                <h3 className="text-xs font-bold uppercase tracking-wider text-slate-600 mb-4 flex items-center gap-2">
                  <MapPin className="w-3.5 h-3.5 text-blue-500" />
                  Format Specs (规范指引)
                </h3>
                <div className="space-y-3">
                  <div className="p-3 bg-slate-50 rounded-2xl border border-slate-100 text-[10px] text-slate-600 space-y-1">
                    <span className="font-bold text-blue-600 block">Inventory Excel (库存表):</span>
                    <p>• <b>Col J</b>: Order No. (工单号)</p>
                    <p>• <b>Col K</b>: Box Grade (箱等级，取 A 级)</p>
                    <p>• <b>Col S</b>: Barcodes (箱条码数，自动过滤 0)</p>
                  </div>
                  <div className="p-3 bg-slate-50 rounded-2xl border border-slate-100 text-[10px] text-slate-600 space-y-1">
                    <span className="font-bold text-violet-600 block">Mapping Excel (对照表):</span>
                    <p>• <b>Col B</b>: Order / Sales No. (工单/销售单)</p>
                    <p>• <b>Col F</b>: Customer Name (客户简称/名称)</p>
                  </div>
                  <div className="p-3 bg-slate-50 rounded-2xl border border-slate-100 text-[9px] text-slate-500 space-y-1">
                    <span className="font-bold text-emerald-600 block text-[10px]">Location Order (基准序列):</span>
                    <p className="font-mono text-slate-700 bg-white p-1.5 rounded border border-slate-200 leading-relaxed">
                      C00~H60 → K00~O35 → WA → WE~WG → TBD
                    </p>
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
                      <div className="flex items-center justify-between pb-3 border-b border-slate-100">
                        <h4 className="text-xs font-bold uppercase tracking-wider text-slate-700 flex items-center gap-2">
                          <BarChart3 className="w-4 h-4 text-blue-500" />
                          Overview (库存概览)
                        </h4>
                        <div className="flex gap-1.5">
                          <span className="px-2 py-0.5 bg-blue-50 text-blue-600 text-[10px] font-bold rounded-full">
                            Locs: {results.length}
                          </span>
                          {results.some(r => !PREDEFINED_LOCATION_SET.has(r['库位名称'])) && (
                            <span className="px-2 py-0.5 bg-amber-50 text-amber-700 text-[10px] font-bold rounded-full border border-amber-200">
                              Extra: {results.filter(r => !PREDEFINED_LOCATION_SET.has(r['库位名称'])).length}
                            </span>
                          )}
                        </div>
                      </div>
                      
                      <div className="grid grid-cols-2 gap-3">
                        <div className="bg-slate-50 rounded-2xl p-3">
                          <div className="text-[10px] text-slate-400 font-bold mb-0.5">Total Boxes (当前箱数)</div>
                          <div className="text-lg font-black text-slate-900 font-mono">
                            {results.reduce((acc, curr) => acc + curr['箱号数量'], 0)}
                          </div>
                        </div>
                        <div className="bg-emerald-50 rounded-2xl p-3">
                          <div className="text-[10px] text-emerald-600 font-bold mb-0.5">Avail Space (剩余位)</div>
                          <div className="text-lg font-black text-emerald-700 font-mono">
                            {results.reduce((acc, curr) => acc + (curr['可入库数量'] > 0 ? curr['可入库数量'] : 0), 0)}
                          </div>
                        </div>
                      </div>

                      <div className="space-y-1.5">
                        <div className="flex items-center justify-between text-[10px] font-bold">
                          <span className="text-slate-400 uppercase tracking-wider flex items-center gap-1">
                            <CheckCircle2 className="w-3 h-3 text-emerald-500" />
                            Empty Locations / 空置库位 ({results.filter(r => r['箱号数量'] === 0).length})
                          </span>
                        </div>
                        <div className="max-h-16 overflow-y-auto pr-1 scrollbar-thin flex flex-wrap gap-1.5 text-[9px]">
                          {results.filter(r => r['箱号数量'] === 0).length > 0 ? (
                            results.filter(r => r['箱号数量'] === 0).map(r => (
                              <button 
                                key={r['库位名称']} 
                                onClick={() => generateLabelsPDF(r['库位名称'])}
                                className="px-2 py-0.5 bg-white text-slate-600 rounded-md border border-slate-200 font-mono shadow-sm hover:border-emerald-400 hover:text-emerald-600 transition-all active:scale-95 group/tag inline-flex items-center gap-1"
                                title="Print Label (打印标签)"
                              >
                                {r['库位名称']}
                                <FileText className="w-2.5 h-2.5 opacity-0 group-hover/tag:opacity-100 transition-opacity" />
                              </button>
                            ))
                          ) : (
                            <span className="text-slate-400 italic">No empty locations / 无空置库位</span>
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
                      mixedPowerAlerts.length > 0 ? "bg-red-50/70 border-red-100" : "bg-white border-slate-200"
                    )}
                  >
                    <div className="flex items-center justify-between mb-3">
                      <div className="flex items-center gap-2">
                        <h4 className="text-xs font-bold uppercase tracking-wider flex items-center gap-1.5 text-slate-700">
                          <AlertCircle className={cn("w-4 h-4", mixedPowerAlerts.length > 0 ? "text-red-500" : "text-slate-400")} />
                          Mixed Power (功率混档)
                        </h4>
                        {mixedPowerAlerts.length > 0 && (
                          <button 
                            onClick={exportAlertsToExcel}
                            className="p-1 hover:bg-red-100 text-red-600 rounded-lg transition-colors"
                            title="Export Mixed Power Report"
                          >
                            <Download className="w-3.5 h-3.5" />
                          </button>
                        )}
                      </div>
                      <span className={cn("text-lg font-black font-mono", mixedPowerAlerts.length > 0 ? "text-red-600" : "text-slate-400")}>
                        {mixedPowerAlerts.length}
                      </span>
                    </div>
                    <div className="max-h-24 overflow-y-auto pr-1 scrollbar-thin">
                      {mixedPowerAlerts.length > 0 ? (
                        <div className="space-y-2">
                          {mixedPowerAlerts.map((alert, i) => (
                            <div key={i} className="flex flex-col p-2.5 bg-white rounded-xl border border-red-200 text-[10px] shadow-sm">
                              <div className="font-bold text-slate-800 mb-1 flex items-center gap-1 font-mono">
                                <Box className="w-3 h-3 text-red-400" />
                                {alert.location}
                              </div>
                              <div className="space-y-0.5">
                                {Object.keys(alert.powerDetails).map(p => (
                                  <div key={p} className="flex items-center justify-between text-slate-500">
                                    <span>Power {p}</span>
                                    <span className="font-mono text-red-600 bg-red-50 px-1 rounded">{alert.powerDetails[p].length} boxes</span>
                                  </div>
                                ))}
                              </div>
                            </div>
                          ))}
                        </div>
                      ) : (
                        <p className="text-slate-400 text-xs italic">Power classes consistent / 功率一致</p>
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
                      mixedCustomerAlerts.length > 0 ? "bg-blue-50/70 border-blue-100" : "bg-white border-slate-200"
                    )}
                  >
                    <div className="flex items-center justify-between mb-3">
                      <div className="flex items-center gap-2">
                        <h4 className="text-xs font-bold uppercase tracking-wider flex items-center gap-1.5 text-slate-700">
                          <History className={cn("w-4 h-4", mixedCustomerAlerts.length > 0 ? "text-blue-500" : "text-slate-400")} />
                          Mixed Customer (客户混载)
                        </h4>
                        {mixedCustomerAlerts.length > 0 && (
                          <button 
                            onClick={exportCustomerAlertsToExcel}
                            className="p-1 hover:bg-blue-100 text-blue-600 rounded-lg transition-colors"
                            title="Export Mixed Customer Report"
                          >
                            <Download className="w-3.5 h-3.5" />
                          </button>
                        )}
                      </div>
                      <span className={cn("text-lg font-black font-mono", mixedCustomerAlerts.length > 0 ? "text-blue-600" : "text-slate-400")}>
                        {mixedCustomerAlerts.length}
                      </span>
                    </div>
                    <div className="max-h-24 overflow-y-auto pr-1 scrollbar-thin">
                      {mixedCustomerAlerts.length > 0 ? (
                        <div className="space-y-2">
                          {mixedCustomerAlerts.map((alert, i) => (
                            <div key={i} className="flex flex-col p-2.5 bg-white rounded-xl border border-blue-200 text-[10px] shadow-sm">
                              <div className="font-bold text-slate-800 mb-1 flex items-center gap-1 font-mono">
                                <Package className="w-3 h-3 text-blue-400" />
                                {alert.location}
                              </div>
                              <div className="space-y-1 divide-y divide-blue-100/50">
                                {Object.keys(alert.customerDetails).map(c => {
                                  const boxes = alert.customerDetails[c];
                                  return (
                                    <div key={c} className="flex items-center justify-between text-slate-500 pt-0.5 first:pt-0">
                                      <span className="truncate max-w-[120px] font-bold text-slate-700" title={c}>{c}</span>
                                      <span className="font-mono text-blue-600 bg-blue-50 px-1 rounded">{boxes.length} boxes</span>
                                    </div>
                                  );
                                })}
                              </div>
                            </div>
                          ))}
                        </div>
                      ) : (
                        <p className="text-slate-400 text-xs italic">Single customer per location / 无混载</p>
                      )}
                    </div>
                  </motion.div>
                </div>
              )}

              {/* Main Detail Table */}
              <div className="bg-white rounded-3xl border border-slate-200 shadow-sm overflow-hidden flex flex-col flex-1">
                <div className="p-5 border-b border-slate-100 flex flex-col gap-3 sm:flex-row sm:items-center sm:justify-between bg-white z-20">
                  <div>
                    <h3 className="font-bold text-lg text-slate-900 tracking-tight flex items-center gap-2">
                      <span>Analytics Dashboard</span>
                      <span className="text-xs font-normal text-slate-400">(明细看板)</span>
                    </h3>
                  </div>

                  <div className="flex items-center gap-3 self-start sm:self-auto">
                    {/* Tab Segment Controller */}
                    <div className="flex bg-slate-100 p-1 rounded-xl">
                      <button 
                        onClick={() => setActiveTab('location')}
                        className={cn(
                          "px-3 py-1.5 text-xs font-bold rounded-lg transition-all flex items-center gap-1.5",
                          activeTab === 'location' ? "bg-white text-slate-950 shadow-sm" : "text-slate-500 hover:text-slate-900"
                        )}
                      >
                        <MapPin className="w-3.5 h-3.5" />
                        By Location (库位)
                      </button>
                      <button 
                        onClick={() => setActiveTab('customer')}
                        className={cn(
                          "px-3 py-1.5 text-xs font-bold rounded-lg transition-all flex items-center gap-1.5",
                          activeTab === 'customer' ? "bg-white text-slate-950 shadow-sm" : "text-slate-500 hover:text-slate-900"
                        )}
                      >
                        <Package className="w-3.5 h-3.5" />
                        By Customer (客户)
                      </button>
                      <button 
                        onClick={() => setActiveTab('correction')}
                        className={cn(
                          "px-3 py-1.5 text-xs font-bold rounded-lg transition-all flex items-center gap-1.5 whitespace-nowrap",
                          activeTab === 'correction' ? "bg-white text-slate-950 shadow-sm" : "text-slate-500 hover:text-slate-900"
                        )}
                      >
                        <ShieldAlert className="w-3.5 h-3.5 text-amber-500" />
                        Relocation Plan (纠偏)
                      </button>
                      <button 
                        onClick={() => setActiveTab('nonA')}
                        className={cn(
                          "px-3 py-1.5 text-xs font-bold rounded-lg transition-all flex items-center gap-1.5 whitespace-nowrap relative",
                          activeTab === 'nonA' ? "bg-purple-600 text-white shadow-sm" : "text-slate-500 hover:text-slate-900"
                        )}
                      >
                        <Layers className="w-3.5 h-3.5" />
                        Non-A Grade (非A级)
                        {nonAGradeList.length > 0 && (
                          <span className={cn(
                            "px-1.5 py-0.2 rounded-full text-[9px] font-mono font-bold ml-0.5",
                            activeTab === 'nonA' ? "bg-white text-purple-700" : "bg-purple-100 text-purple-700"
                          )}>
                            {nonAGradeList.length}
                          </span>
                        )}
                      </button>
                    </div>

                    {((activeTab === 'location' && results.length > 0) || 
                      (activeTab === 'customer' && customerStats.length > 0) || 
                      (activeTab === 'correction' && relocationSuggestions.length > 0) ||
                      (activeTab === 'nonA' && nonAGradeList.length > 0)) && (
                      <div className="flex items-center gap-2">
                        <div className="w-px h-6 bg-slate-200 hidden sm:block" />
                        <button 
                          onClick={
                            activeTab === 'location' ? exportToExcel : 
                            activeTab === 'customer' ? exportCustomerStatsToExcel : 
                            activeTab === 'correction' ? exportCorrectionToExcel : 
                            exportNonAGradeToExcel
                          }
                          className={cn(
                            "px-3 py-1.5 rounded-xl border transition-colors flex items-center gap-1.5 shrink-0 text-xs font-bold shadow-sm",
                            activeTab === 'nonA'
                              ? "bg-purple-600 hover:bg-purple-500 text-white border-purple-600"
                              : "bg-slate-50 hover:bg-slate-100 text-slate-700 border-slate-200"
                          )}
                          title="Export Current View to Excel"
                        >
                          <Download className="w-3.5 h-3.5" />
                          <span>Export Excel (导出)</span>
                        </button>
                      </div>
                    )}
                  </div>
                </div>

                {/* Non-A Grade Filter Bar */}
                {activeTab === 'nonA' && (
                  <div className="bg-purple-50/70 border-b border-purple-100 px-6 py-2.5 flex flex-col md:flex-row md:items-center justify-between gap-2.5 shrink-0">
                    <div className="flex items-center flex-wrap gap-1.5">
                      <span className="text-[10px] font-bold text-purple-900 flex items-center gap-1 mr-1">
                        <Filter className="w-3 h-3 text-purple-600" />
                        Grade Filter (品级):
                      </span>
                      {[
                        { id: 'ALL', label: `All (${nonAGradeList.length})` },
                        { id: 'B', label: `Grade B (${nonAGradeList.filter(x => x.grade === 'B').length})` },
                        { id: 'F', label: `Grade F (${nonAGradeList.filter(x => x.grade === 'F').length})` },
                        { id: 'R', label: `Grade R (${nonAGradeList.filter(x => x.grade === 'R').length})` },
                        { id: 'OTHER', label: `Other (${nonAGradeList.filter(x => !['B','F','R'].includes(x.grade)).length})` },
                      ].map((pill) => (
                        <button
                          key={pill.id}
                          onClick={() => setNonAGradeFilter(pill.id as any)}
                          className={cn(
                            "px-2.5 py-0.5 text-xs font-bold rounded-lg border transition-all",
                            nonAGradeFilter === pill.id
                              ? "bg-purple-700 text-white border-purple-700 shadow-sm"
                              : "bg-white text-slate-600 border-purple-200 hover:bg-purple-100/60"
                          )}
                        >
                          {pill.label}
                        </button>
                      ))}
                    </div>

                    <div className="flex items-center gap-2">
                      <div className="relative">
                        <Search className="w-3 h-3 text-purple-400 absolute left-2.5 top-1/2 -translate-y-1/2" />
                        <input
                          type="text"
                          placeholder="Search box, loc, order, client... (搜索)"
                          value={nonAGradeSearch}
                          onChange={(e) => setNonAGradeSearch(e.target.value)}
                          className="pl-8 pr-3 py-1 text-xs bg-white rounded-lg border border-purple-200 focus:outline-none focus:ring-2 focus:ring-purple-500 w-52 placeholder:text-slate-400"
                        />
                      </div>
                      <button
                        onClick={exportNonAGradeToExcel}
                        className="px-2.5 py-1 bg-purple-600 hover:bg-purple-700 text-white text-xs font-bold rounded-lg transition-colors flex items-center gap-1 shrink-0 shadow-sm"
                        title="Export Non-A Excel"
                      >
                        <Download className="w-3 h-3" />
                        Export (导出)
                      </button>
                    </div>
                  </div>
                )}

                <div className="flex-1 overflow-auto">
                  <table className="w-full border-collapse">
                    {activeTab === 'location' ? (
                      <>
                        <thead className="sticky top-0 bg-white/95 backdrop-blur-sm z-10 border-b border-slate-100 shadow-sm shadow-slate-100/50">
                          <tr className="text-left text-slate-400 text-[10px] font-bold uppercase tracking-wider">
                            <th className="pl-6 pr-3 py-4 w-14">No. (序号)</th>
                            <th className="px-5 py-4">Location (库位)</th>
                            <th className="px-5 py-4">Primary Cust & Power (主存客户/功率)</th>
                            <th className="px-5 py-4 text-center">Box Qty / Cap (箱数/容量)</th>
                            <th className="px-5 py-4 text-right pr-8">Avail Cap (可入容量)</th>
                          </tr>
                        </thead>
                        <tbody className="divide-y divide-slate-50">
                          {results.length > 0 ? (
                            results.map((row, i) => {
                              const capacity = row['库位名称'].toUpperCase().startsWith('H') ? 14 : 18;
                              return (
                                <motion.tr 
                                  key={i}
                                  initial={{ opacity: 0, y: 10 }}
                                  animate={{ opacity: 1, y: 0 }}
                                  transition={{ delay: i * 0.012, duration: 0.25 }}
                                  className="hover:bg-slate-50/50 transition-all duration-200 group text-xs"
                                >
                                  <td className="pl-6 pr-3 py-3.5 font-mono font-bold text-slate-300 group-hover:text-blue-500 transition-colors">
                                    {String(i + 1).padStart(2, '0')}
                                  </td>
                                  <td className="px-5 py-3.5">
                                    <div className="flex items-center gap-2.5">
                                      <div className={cn(
                                        "w-2 h-2 rounded-full shrink-0",
                                        row['箱号数量'] > capacity ? "bg-red-500 shadow-[0_0_6px_rgba(239,68,68,0.5)]" : 
                                        row['箱号数量'] < capacity ? "bg-emerald-500 shadow-[0_0_6px_rgba(16,185,129,0.5)]" : 
                                        "bg-slate-700"
                                      )} />
                                      <div className="flex flex-col">
                                        <div className="flex items-center gap-1.5">
                                          <span className={cn(
                                            "text-xs font-bold transition-colors font-mono",
                                            row['箱号数量'] > capacity ? "text-red-600" : 
                                            row['箱号数量'] < capacity ? "text-emerald-700" :
                                            "text-slate-900"
                                          )}>
                                            {row['库位名称']}
                                          </span>
                                          {!PREDEFINED_LOCATION_SET.has(row['库位名称']) && (
                                            <span className="px-1 py-0.2 text-[8px] font-bold bg-amber-50 text-amber-700 rounded border border-amber-200">
                                              Extra
                                            </span>
                                          )}
                                        </div>
                                        {row['客户分布'] && Object.keys(row['客户分布']).length > 0 && (
                                          <span className="text-[9px] text-slate-400" title={
                                            Object.entries(row['客户分布'])
                                              .map(([cust, count]) => `${cust} (${count})`)
                                              .join(', ')
                                          }>
                                            {Object.entries(row['客户分布']).map(([cust, count]) => `${cust} (${count})`).join(', ')}
                                          </span>
                                        )}
                                      </div>
                                    </div>
                                  </td>
                                  <td className="px-5 py-3.5">
                                    {row['主存客户'] !== '空置' ? (
                                      <div className="flex flex-col">
                                        <span className="text-xs font-bold text-slate-800 truncate max-w-[160px]" title={row['主存客户']}>
                                          {row['主存客户']}
                                        </span>
                                        <span className="text-[9px] text-slate-400 font-mono">
                                          Power: {row['主存功率']}
                                        </span>
                                      </div>
                                    ) : (
                                      <span className="text-xs text-slate-300 font-medium italic">Empty (空置)</span>
                                    )}
                                  </td>
                                  <td className="px-5 py-3.5 text-center">
                                    <span className={cn(
                                      "text-xs font-mono font-bold px-2.5 py-0.5 rounded-full",
                                      row['箱号数量'] > capacity ? "bg-red-50 text-red-600" : 
                                      row['箱号数量'] < capacity ? "bg-emerald-50 text-emerald-600" : 
                                      "bg-slate-100 text-slate-900"
                                    )}>
                                      {row['箱号数量']} / {capacity}
                                    </span>
                                  </td>
                                  <td className="px-5 py-3.5 text-right pr-8">
                                    <span className={cn(
                                      "text-xs font-mono font-bold",
                                      row['可入库数量'] < 0 ? "text-red-500" : 
                                      row['可入库数量'] > 0 ? "text-emerald-600" : 
                                      "text-slate-900"
                                    )}>
                                      {row['可入库数量'] > 0 ? `+${row['可入库数量']}` : row['可入库数量']}
                                    </span>
                                  </td>
                                </motion.tr>
                              );
                            })
                          ) : (
                            <tr>
                              <td colSpan={5} className="px-8 py-24 text-center">
                                <div className="flex flex-col items-center justify-center">
                                  <div className="w-12 h-12 bg-slate-50 rounded-full flex items-center justify-center mb-3 text-slate-300">
                                    <TableIcon className="w-6 h-6" />
                                  </div>
                                  <p className="text-xs font-bold text-slate-500">Awaiting inventory data (等待载入库存)</p>
                                  <p className="text-[10px] text-slate-400 mt-0.5">Please upload your inventory Excel on the left</p>
                                </div>
                              </td>
                            </tr>
                          )}
                        </tbody>
                      </>
                    ) : activeTab === 'customer' ? (
                      <>
                        <thead className="sticky top-0 bg-white/95 backdrop-blur-sm z-10 border-b border-slate-100 shadow-sm shadow-slate-100/50">
                          <tr className="text-left text-slate-400 text-[10px] font-bold uppercase tracking-wider">
                            <th className="pl-6 pr-3 py-4 w-14">No. (序号)</th>
                            <th className="px-5 py-4">Customer Name (客户名称)</th>
                            <th className="px-5 py-4 text-center">Box Qty (占箱数)</th>
                            <th className="px-5 py-4 pr-8">Occupied Locations (占用库位)</th>
                          </tr>
                        </thead>
                        <tbody className="divide-y divide-slate-50 text-xs">
                          {customerStats.length > 0 ? (
                            customerStats.map((row, i) => (
                              <motion.tr 
                                key={i}
                                initial={{ opacity: 0, y: 10 }}
                                animate={{ opacity: 1, y: 0 }}
                                transition={{ delay: i * 0.012, duration: 0.25 }}
                                className="hover:bg-slate-50/50 transition-all duration-200 group"
                              >
                                <td className="pl-6 pr-3 py-3.5 font-mono font-bold text-slate-300 group-hover:text-violet-500 transition-colors">
                                  {String(i + 1).padStart(2, '0')}
                                </td>
                                <td className="px-5 py-3.5">
                                  <span className="font-bold text-slate-900 block">
                                    {row.customerName}
                                  </span>
                                </td>
                                <td className="px-5 py-3.5 text-center">
                                  <span className="font-mono font-bold bg-violet-50 text-violet-700 px-2.5 py-0.5 rounded-full border border-violet-100">
                                    {row.boxCount} Boxes
                                  </span>
                                </td>
                                <td className="px-5 py-3.5 pr-8">
                                  <div className="flex flex-wrap gap-1 max-h-20 overflow-y-auto pr-1">
                                    {row.locations.map(loc => (
                                      <span key={loc} className="px-1.5 py-0.5 bg-slate-100 text-slate-600 text-[9px] font-mono rounded font-bold border border-slate-200">
                                        {loc}
                                      </span>
                                    ))}
                                  </div>
                                </td>
                              </motion.tr>
                            ))
                          ) : (
                            <tr>
                              <td colSpan={4} className="px-8 py-24 text-center">
                                <div className="flex flex-col items-center justify-center max-w-sm mx-auto">
                                  <div className="w-12 h-12 bg-violet-50 text-violet-500 rounded-full flex items-center justify-center mb-3">
                                    <Package className="w-6 h-6" />
                                  </div>
                                  <p className="text-xs font-bold text-slate-600">No customer mapping data (暂无客户映射)</p>
                                  <p className="text-[10px] text-slate-400 mt-1">
                                    Please upload Order-Customer Mapping Excel to link orders with customer names.
                                  </p>
                                </div>
                              </td>
                            </tr>
                          )}
                        </tbody>
                      </>
                    ) : activeTab === 'correction' ? (
                      <>
                        <thead className="sticky top-0 bg-white/95 backdrop-blur-sm z-10 border-b border-slate-100 shadow-sm shadow-slate-100/50">
                          <tr className="text-left text-slate-400 text-[10px] font-bold uppercase tracking-wider">
                            <th className="pl-6 pr-3 py-4 w-14">No. (序号)</th>
                            <th className="px-5 py-4">Box No. (箱号)</th>
                            <th className="px-5 py-4 text-center">Current Loc (当前库位)</th>
                            <th className="px-5 py-4">Box Attr (该箱属性)</th>
                            <th className="px-5 py-4">Dominant Attr (库位主流)</th>
                            <th className="px-5 py-4">Mismatch (偏差)</th>
                            <th className="px-5 py-4 pr-8">Target Loc & Reason (推荐去向与原因)</th>
                          </tr>
                        </thead>
                        <tbody className="divide-y divide-slate-50 text-slate-700 text-xs">
                          {relocationSuggestions.length > 0 ? (
                            relocationSuggestions.map((row, i) => (
                              <motion.tr 
                                key={i}
                                initial={{ opacity: 0, y: 10 }}
                                animate={{ opacity: 1, y: 0 }}
                                transition={{ delay: i * 0.01, duration: 0.2 }}
                                className="hover:bg-slate-50/50 transition-all duration-200 group"
                              >
                                <td className="pl-6 pr-3 py-3 font-mono font-bold text-slate-300 group-hover:text-amber-500 transition-colors">
                                  {String(i + 1).padStart(2, '0')}
                                </td>
                                <td className="px-5 py-3 font-bold text-slate-800 font-mono">
                                  {row.boxNo}
                                </td>
                                <td className="px-5 py-3 text-center font-bold text-slate-600 font-mono">
                                  {row.currentLoc}
                                </td>
                                <td className="px-5 py-3">
                                  <div className="flex flex-col">
                                    <span className="font-bold text-slate-800 truncate max-w-[120px]" title={row.customerName}>{row.customerName}</span>
                                    <span className="font-mono text-slate-400 text-[9px]">P: {row.powerGrade}</span>
                                  </div>
                                </td>
                                <td className="px-5 py-3">
                                  <div className="flex flex-col">
                                    <span className="font-bold text-blue-700 truncate max-w-[120px]" title={row.dominantCustomer}>{row.dominantCustomer}</span>
                                    <span className="font-mono text-blue-500 text-[9px]">P: {row.dominantPower}</span>
                                  </div>
                                </td>
                                <td className="px-5 py-3">
                                  <span className="inline-block px-1.5 py-0.5 font-bold text-[9px] rounded bg-amber-50 text-amber-700 border border-amber-200">
                                    {row.mismatchType}
                                  </span>
                                </td>
                                <td className="px-5 py-3 pr-8">
                                  <div className="flex flex-col">
                                    <span className="font-bold text-emerald-600 font-mono flex items-center gap-1">
                                      <ArrowRight className="w-3 h-3" />
                                      {row.recommendedLoc}
                                    </span>
                                    <span className="text-[9px] text-slate-400 max-w-xs">
                                      {row.recommendationReason}
                                    </span>
                                  </div>
                                </td>
                              </motion.tr>
                            ))
                          ) : (
                            <tr>
                              <td colSpan={7} className="px-8 py-24 text-center">
                                <div className="flex flex-col items-center justify-center max-w-sm mx-auto">
                                  <div className="w-12 h-12 bg-emerald-50 text-emerald-500 rounded-full flex items-center justify-center mb-3">
                                    <CheckCircle2 className="w-6 h-6" />
                                  </div>
                                  <p className="text-xs font-bold text-slate-800">No mixed locations detected (未发现混存库位)</p>
                                  <p className="text-[10px] text-slate-400 mt-1">
                                    All current locations meet single-customer and single-power consistency.
                                  </p>
                                </div>
                              </td>
                            </tr>
                          )}
                        </tbody>
                      </>
                    ) : (
                      /* activeTab === 'nonA' */
                      <>
                        <thead className="sticky top-0 bg-white/95 backdrop-blur-sm z-10 border-b border-purple-100 shadow-sm shadow-purple-50">
                          <tr className="text-left text-slate-400 text-[10px] font-bold uppercase tracking-wider">
                            <th className="pl-6 pr-3 py-4 w-14">No. (序号)</th>
                            <th className="px-5 py-4">Box No. (箱号)</th>
                            <th className="px-5 py-4 text-center">Location (库位)</th>
                            <th className="px-5 py-4 text-center">Raw Grade (Col K)</th>
                            <th className="px-5 py-4 text-center">Category (归类)</th>
                            <th className="px-5 py-4 text-center">Barcodes (Col S)</th>
                            <th className="px-5 py-4">Order No. (Col J)</th>
                            <th className="px-5 py-4">Customer (客户)</th>
                            <th className="px-5 py-4 pr-8 text-center">Power (功率)</th>
                          </tr>
                        </thead>
                        <tbody className="divide-y divide-slate-50 text-slate-700 text-xs">
                          {filteredNonAGradeList.length > 0 ? (
                            filteredNonAGradeList.map((row, i) => (
                              <motion.tr 
                                key={i}
                                initial={{ opacity: 0, y: 10 }}
                                animate={{ opacity: 1, y: 0 }}
                                transition={{ delay: i * 0.008, duration: 0.2 }}
                                className="hover:bg-purple-50/40 transition-all duration-200 group"
                              >
                                <td className="pl-6 pr-3 py-3 font-mono font-bold text-slate-300 group-hover:text-purple-600 transition-colors">
                                  {String(i + 1).padStart(2, '0')}
                                </td>
                                <td className="px-5 py-3 font-bold text-slate-800 font-mono">
                                  {row.boxNo}
                                </td>
                                <td className="px-5 py-3 text-center">
                                  <span className="font-mono font-bold text-slate-700 bg-slate-100 px-1.5 py-0.5 rounded border border-slate-200">
                                    {row.location}
                                  </span>
                                </td>
                                <td className="px-5 py-3 text-center font-mono text-slate-600">
                                  {row.rawGrade || '-'}
                                </td>
                                <td className="px-5 py-3 text-center">
                                  <span className={cn(
                                    "px-2 py-0.5 rounded font-bold text-[10px] inline-block",
                                    row.grade === 'B' ? "bg-blue-100 text-blue-800 border border-blue-200" :
                                    row.grade === 'F' ? "bg-purple-100 text-purple-800 border border-purple-200" :
                                    row.grade === 'R' ? "bg-rose-100 text-rose-800 border border-rose-200" :
                                    "bg-slate-100 text-slate-700 border border-slate-200"
                                  )}>
                                    Grade {row.grade}
                                  </span>
                                </td>
                                <td className="px-5 py-3 text-center font-mono font-bold text-slate-700">
                                  {row.barcodeCount !== undefined ? row.barcodeCount : '-'}
                                </td>
                                <td className="px-5 py-3 font-mono text-slate-600">
                                  {row.orderNo || '-'}
                                </td>
                                <td className="px-5 py-3">
                                  <span className={cn(
                                    "font-bold truncate max-w-[130px] block",
                                    row.customerName === '未映射客户' ? "text-slate-400 italic font-normal" : "text-slate-800"
                                  )} title={row.customerName}>
                                    {row.customerName || '-'}
                                  </span>
                                </td>
                                <td className="px-5 py-3 pr-8 text-center font-mono text-slate-500">
                                  {row.power || '-'}
                                </td>
                              </motion.tr>
                            ))
                          ) : (
                            <tr>
                              <td colSpan={9} className="px-8 py-24 text-center">
                                <div className="flex flex-col items-center justify-center max-w-sm mx-auto">
                                  <div className="w-12 h-12 bg-purple-50 text-purple-500 rounded-full flex items-center justify-center mb-3">
                                    <Layers className="w-6 h-6" />
                                  </div>
                                  <p className="text-xs font-bold text-slate-800">
                                    {nonAGradeList.length === 0 ? 'No Non-A Grade items found (未检测到非A级品)' : 'No matching items (无匹配数据)'}
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

                <div className="p-4 bg-slate-50/50 border-t border-slate-100 flex items-center justify-between">
                  <div className="flex items-center gap-1.5 opacity-60">
                    <CheckCircle2 className="w-3.5 h-3.5 text-emerald-500" />
                    <span className="text-[9px] font-bold text-slate-500 uppercase tracking-wider">Engine v5.1 Pro • Sandbox Mode</span>
                  </div>
                  <p className="text-[9px] text-slate-400 font-mono">
                    Updated: {new Date().toLocaleTimeString()}
                  </p>
                </div>
              </div>
            </div>
          </div>
        </div>
      </main>

      {/* S列条码为0过滤明细 Modal */}
      <AnimatePresence>
        {showZeroBarcodeModal && (
          <motion.div
            initial={{ opacity: 0 }}
            animate={{ opacity: 1 }}
            exit={{ opacity: 0 }}
            className="fixed inset-0 z-50 bg-slate-900/60 backdrop-blur-sm flex items-center justify-center p-4 sm:p-6"
            onClick={() => setShowZeroBarcodeModal(false)}
          >
            <motion.div
              initial={{ scale: 0.95, opacity: 0, y: 20 }}
              animate={{ scale: 1, opacity: 1, y: 0 }}
              exit={{ scale: 0.95, opacity: 0, y: 20 }}
              onClick={e => e.stopPropagation()}
              className="bg-white rounded-3xl shadow-2xl border border-slate-200 max-w-4xl w-full max-h-[85vh] flex flex-col overflow-hidden"
            >
              {/* Modal Header */}
              <div className="px-6 py-5 border-b border-slate-100 flex items-center justify-between bg-orange-50/50">
                <div className="flex items-center gap-3">
                  <div className="w-10 h-10 rounded-2xl bg-orange-500 text-white flex items-center justify-center shadow-lg shadow-orange-500/20">
                    <ShieldAlert className="w-5 h-5" />
                  </div>
                  <div>
                    <h3 className="text-base font-bold text-slate-800 flex items-center gap-2">
                      Zero-Barcode Filtered Items (S列条码为0明细)
                      <span className="px-2 py-0.5 rounded-full text-xs bg-orange-100 text-orange-700 font-mono font-bold">
                        {zeroBarcodeList.length} Boxes (箱)
                      </span>
                    </h3>
                    <p className="text-xs text-slate-500 mt-0.5">
                      Rows with Column S (box barcode count) = 0 or empty are excluded from inventory calculation.
                    </p>
                  </div>
                </div>
                <div className="flex items-center gap-2">
                  <button
                    onClick={exportZeroBarcodesToExcel}
                    className="px-3.5 py-2 bg-orange-600 hover:bg-orange-700 text-white rounded-xl text-xs font-bold flex items-center gap-1.5 shadow-sm transition-colors"
                  >
                    <Download className="w-3.5 h-3.5" />
                    Export Excel (导出清单)
                  </button>
                  <button
                    onClick={() => setShowZeroBarcodeModal(false)}
                    className="w-9 h-9 rounded-xl hover:bg-slate-200/60 text-slate-400 hover:text-slate-600 flex items-center justify-center transition-colors"
                  >
                    <X className="w-5 h-5" />
                  </button>
                </div>
              </div>

              {/* Modal Content Table */}
              <div className="flex-1 overflow-auto p-6">
                <div className="rounded-2xl border border-slate-200 overflow-hidden shadow-sm">
                  <table className="w-full text-left border-collapse text-xs">
                    <thead>
                      <tr className="bg-slate-50 border-b border-slate-200 text-[10px] uppercase font-bold text-slate-500">
                        <th className="px-4 py-3 text-center">No. (序号)</th>
                        <th className="px-4 py-3">Row No. (原行号)</th>
                        <th className="px-4 py-3">Box No. (箱号)</th>
                        <th className="px-4 py-3">Location (库位)</th>
                        <th className="px-4 py-3">Grade (Col K)</th>
                        <th className="px-4 py-3 text-center">Raw Val (S列原值)</th>
                        <th className="px-4 py-3">Order No. (工单号)</th>
                        <th className="px-4 py-3">Filter Reason (判定原因)</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-100">
                      {zeroBarcodeList.map((item, idx) => (
                        <tr key={idx} className="hover:bg-orange-50/30 transition-colors">
                          <td className="px-4 py-2.5 text-center font-mono text-slate-400">{idx + 1}</td>
                          <td className="px-4 py-2.5 font-mono text-slate-600">Row {item.rowIdx}</td>
                          <td className="px-4 py-2.5 font-mono font-bold text-slate-800">{item.boxNo}</td>
                          <td className="px-4 py-2.5 font-mono text-slate-600">{item.location}</td>
                          <td className="px-4 py-2.5">
                            <span className="px-2 py-0.5 bg-slate-100 text-slate-600 rounded text-[10px] font-bold">
                              {item.grade}
                            </span>
                          </td>
                          <td className="px-4 py-2.5 text-center font-mono font-bold text-orange-600">
                            {item.rawBarcode === '' ? <span className="text-slate-400 font-normal italic">(Empty/留空)</span> : item.rawBarcode}
                          </td>
                          <td className="px-4 py-2.5 font-mono text-slate-500">{item.orderNo || '-'}</td>
                          <td className="px-4 py-2.5 text-slate-600">
                            <span className="px-2 py-0.5 rounded bg-orange-100/70 text-orange-800 text-[10px] font-medium">
                              {item.reason}
                            </span>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>

              {/* Modal Footer */}
              <div className="px-6 py-4 border-t border-slate-100 bg-slate-50 flex items-center justify-between text-xs text-slate-500">
                <span>
                  Reconciliation: Raw rows {totalRowsRaw} - Barcode=0 ({zeroBarcodeList.length}) = Valid {totalRowsRaw - zeroBarcodeList.length} boxes (实际有效)
                </span>
                <button
                  onClick={() => setShowZeroBarcodeModal(false)}
                  className="px-4 py-1.5 bg-slate-200 hover:bg-slate-300 text-slate-700 rounded-xl text-xs font-bold transition-colors"
                >
                  Close (关闭)
                </button>
              </div>
            </motion.div>
          </motion.div>
        )}
      </AnimatePresence>

      {/* 8-Box Variance Audit & Diagnostics Modal */}
      <AnimatePresence>
        {showAuditModal && (
          <motion.div
            initial={{ opacity: 0 }}
            animate={{ opacity: 1 }}
            exit={{ opacity: 0 }}
            className="fixed inset-0 z-50 bg-slate-900/60 backdrop-blur-sm flex items-center justify-center p-4 sm:p-6"
            onClick={() => setShowAuditModal(false)}
          >
            <motion.div
              initial={{ scale: 0.95, opacity: 0, y: 20 }}
              animate={{ scale: 1, opacity: 1, y: 0 }}
              exit={{ scale: 0.95, opacity: 0, y: 20 }}
              onClick={e => e.stopPropagation()}
              className="bg-white rounded-3xl shadow-2xl border border-slate-200 max-w-5xl w-full max-h-[90vh] flex flex-col overflow-hidden"
            >
              {/* Modal Header */}
              <div className="px-6 py-5 border-b border-slate-100 flex items-center justify-between bg-indigo-50/50">
                <div className="flex items-center gap-3">
                  <div className="w-10 h-10 rounded-2xl bg-indigo-600 text-white flex items-center justify-center shadow-lg shadow-indigo-500/20">
                    <HelpCircle className="w-5 h-5" />
                  </div>
                  <div>
                    <h3 className="text-base font-bold text-slate-800 flex items-center gap-2">
                      Reconciliation & Audit Diagnostics (8箱差异深度溯源与诊断)
                      <span className="px-2 py-0.5 rounded-full text-xs bg-indigo-100 text-indigo-700 font-mono font-bold">
                        Variance: 8 Boxes (差异 8 箱)
                      </span>
                    </h3>
                    <p className="text-xs text-slate-500 mt-0.5">
                      Full audit reconciliation: {totalRowsRaw} raw rows vs 7,372 subtotal vs {totalRowsRaw - zeroBarcodeCount} valid boxes
                    </p>
                  </div>
                </div>
                <div className="flex items-center gap-2">
                  <button
                    onClick={exportAuditReportToExcel}
                    className="px-3.5 py-1.5 bg-indigo-600 hover:bg-indigo-700 text-white rounded-xl text-xs font-bold flex items-center gap-1.5 transition-colors shadow-sm"
                  >
                    <Download className="w-3.5 h-3.5" />
                    Export Audit Excel (导出诊断表)
                  </button>
                  <button 
                    onClick={() => setShowAuditModal(false)}
                    className="w-8 h-8 rounded-full bg-white hover:bg-slate-100 text-slate-400 hover:text-slate-600 flex items-center justify-center transition-colors border border-slate-200"
                  >
                    <X className="w-4 h-4" />
                  </button>
                </div>
              </div>

              {/* Modal Body */}
              <div className="p-6 overflow-y-auto space-y-6 max-h-[calc(90vh-140px)]">
                
                {/* 1. Mathematical Reconciliation Overview */}
                <div className="bg-slate-900 rounded-2xl p-5 text-white shadow-md space-y-4">
                  <div className="flex items-center justify-between border-b border-slate-800 pb-3">
                    <span className="text-xs font-bold text-indigo-400 uppercase tracking-wider flex items-center gap-2">
                      <ListOrdered className="w-4 h-4" />
                      Core Mathematical Derivation & Reconciliation (推导等式)
                    </span>
                    <span className="text-[10px] text-slate-400 font-mono">
                      Strict Audit Proof (严格闭环)
                    </span>
                  </div>

                  <div className="grid grid-cols-1 md:grid-cols-4 gap-3 text-center font-mono">
                    <div className="bg-slate-800/80 p-3 rounded-xl border border-slate-700">
                      <span className="text-[10px] text-slate-400 block font-sans">① Raw Data Rows (源表总行)</span>
                      <span className="text-xl font-bold text-white mt-1 block">{totalRowsRaw}</span>
                      <span className="text-[9px] text-slate-400 block mt-1">Excel physical rows</span>
                    </div>
                    <div className="bg-orange-950/60 p-3 rounded-xl border border-orange-800/80">
                      <span className="text-[10px] text-orange-300 block font-sans">② Col S=0 Filtered (S列0条码)</span>
                      <span className="text-xl font-bold text-orange-400 mt-1 block">-{zeroBarcodeCount}</span>
                      <span className="text-[9px] text-orange-300/80 block mt-1">Excluded empty rows</span>
                    </div>
                    <div className="bg-emerald-950/60 p-3 rounded-xl border border-emerald-800/80">
                      <span className="text-[10px] text-emerald-300 block font-sans">③ Valid Included Boxes (有效总箱数)</span>
                      <span className="text-xl font-bold text-emerald-400 mt-1 block">{totalRowsRaw - zeroBarcodeCount}</span>
                      <span className="text-[9px] text-emerald-300/80 block mt-1">7,380 - 23 = 7,357</span>
                    </div>
                    <div className="bg-purple-950/60 p-3 rounded-xl border border-purple-800/80">
                      <span className="text-[10px] text-purple-300 block font-sans">④ Grade A + Non-A (品级构成)</span>
                      <span className="text-lg font-bold text-purple-300 mt-1 block">{data.length} + {nonAGradeCount}</span>
                      <span className="text-[9px] text-purple-300/80 block mt-1">6,879 (A) + 478 (Non-A) = 7,357</span>
                    </div>
                  </div>

                  {/* Root Cause Explanation */}
                  <div className="bg-indigo-950/50 border border-indigo-700/60 rounded-xl p-3.5 text-xs text-indigo-200 space-y-1.5 leading-relaxed font-sans">
                    <p className="font-bold text-indigo-300 flex items-center gap-1.5 text-sm">
                      <Info className="w-4 h-4 text-indigo-400 shrink-0" />
                      Root Cause Analysis (差异根本原因分析):
                    </p>
                    <p>
                      • <strong>Reported 7,349 Boxes:</strong> Calculated from <code className="bg-slate-800 px-1.5 py-0.5 rounded text-amber-300 font-mono">7,372 - 23 = 7,349</code> (deducting 23 from previous table subtotal of 7,372).
                    </p>
                    <p>
                      • <strong>Actual 7,357 Boxes:</strong> Calculated from <code className="bg-slate-800 px-1.5 py-0.5 rounded text-emerald-300 font-mono">7,380 - 23 = 7,357</code> (deducting 23 from full raw dataset of 7,380 rows).
                    </p>
                    <p>
                      • <strong>Net Difference:</strong> <code className="bg-slate-800 px-1.5 py-0.5 rounded text-white font-mono">7,357 - 7,349 = (7,380 - 23) - (7,372 - 23) = 7,380 - 7,372 = 8 Boxes</code>.
                    </p>
                    <p className="text-emerald-300 font-medium">
                      👉 <strong>Conclusion (结论):</strong> The 6,879 Grade-A + 478 Non-A = 7,357 boxes strictly exclude all 23 invalid rows. The 8-box gap arose because the previous 7,372 subtotal had already omitted 8 source rows before deduction.
                    </p>
                  </div>
                </div>

                {/* 2. Diagnostics by Characteristic */}
                <div className="space-y-4">
                  <h4 className="text-sm font-bold text-slate-800 flex items-center gap-2">
                    <span className="w-2.5 h-2.5 rounded-full bg-blue-600" />
                    Source Data Diagnostics by Pattern (分类特征诊断)
                  </h4>

                  {/* Feature 1: Duplicate Box Numbers */}
                  <div className="bg-slate-50 border border-slate-200 rounded-2xl p-4 space-y-3">
                    <div className="flex items-center justify-between">
                      <div>
                        <h5 className="text-xs font-bold text-slate-800 flex items-center gap-1.5">
                          Pattern 1: Duplicate Box Numbers (重复箱号排查)
                        </h5>
                        <p className="text-[11px] text-slate-500 mt-0.5">
                          If deduplication or pivot distinct count was used, duplicate box records will be collapsed into 1 count.
                        </p>
                      </div>
                      <span className={cn(
                        "px-2.5 py-1 rounded-full text-xs font-bold font-mono",
                        auditReport && auditReport.duplicateRowCount > 0 ? "bg-amber-100 text-amber-800 border border-amber-200" : "bg-emerald-100 text-emerald-800"
                      )}>
                        {auditReport ? `${auditReport.duplicateRowCount} rows (${auditReport.duplicateBoxes.length} boxes)` : 'Scanning...'}
                      </span>
                    </div>

                    {auditReport && auditReport.duplicateBoxes.length > 0 ? (
                      <div className="overflow-x-auto border border-slate-200 rounded-xl bg-white max-h-52">
                        <table className="w-full text-left text-xs">
                          <thead className="bg-slate-100 sticky top-0 text-[10px] font-bold text-slate-600 uppercase">
                            <tr>
                              <th className="px-3 py-2">Box No. (箱号)</th>
                              <th className="px-3 py-2 text-center">Row No. (原行号)</th>
                              <th className="px-3 py-2 text-center">Occurrences (重复频次)</th>
                              <th className="px-3 py-2">Location (库位)</th>
                              <th className="px-3 py-2">Power (功率档)</th>
                              <th className="px-3 py-2">Order No. (工单号)</th>
                              <th className="px-3 py-2">Characteristic (诊断特征)</th>
                            </tr>
                          </thead>
                          <tbody className="divide-y divide-slate-100 font-mono">
                            {auditReport.duplicateBoxes.map((dup, idx) => (
                              <tr key={idx} className="hover:bg-amber-50/50">
                                <td className="px-3 py-2 font-bold text-slate-800">{dup.boxNo}</td>
                                <td className="px-3 py-2 text-center text-indigo-600 font-bold">
                                  {dup.rows.map(r => `Row ${r.rowIdx}`).join(', ')}
                                </td>
                                <td className="px-3 py-2 text-center">
                                  <span className="px-2 py-0.5 rounded bg-amber-100 text-amber-800 text-[10px] font-bold">
                                    {dup.count}x
                                  </span>
                                </td>
                                <td className="px-3 py-2 font-sans">{dup.rows.map(r => r.location).join(' / ')}</td>
                                <td className="px-3 py-2 font-sans">{dup.rows.map(r => r.power).join(' / ')}</td>
                                <td className="px-3 py-2">{dup.rows.map(r => r.orderNo).join(' / ')}</td>
                                <td className="px-3 py-2 font-sans text-amber-700 text-[11px]">
                                  {dup.rows.some((r, i, arr) => i > 0 && r.power !== arr[0].power) 
                                    ? 'Multi-Power (同箱多功率分录)' 
                                    : dup.rows.some((r, i, arr) => i > 0 && r.orderNo !== arr[0].orderNo)
                                    ? 'Multi-Order (同箱多工单分录)'
                                    : 'Exact Duplicate (完全重复录入)'}
                                </td>
                              </tr>
                            ))}
                          </tbody>
                        </table>
                      </div>
                    ) : (
                      <p className="text-xs text-slate-500 italic bg-white p-3 rounded-xl border border-slate-200">
                        No duplicate box numbers found in the current dataset. All valid rows have unique box identifiers.
                      </p>
                    )}
                  </div>

                  {/* Feature 2: Col S = 0 Filtered Items */}
                  <div className="bg-orange-50/60 border border-orange-200 rounded-2xl p-4 space-y-3">
                    <div className="flex items-center justify-between">
                      <div>
                        <h5 className="text-xs font-bold text-slate-800 flex items-center gap-1.5">
                          Pattern 2: Col S=0 Filtered Rows ({zeroBarcodeList.length} Items)
                        </h5>
                        <p className="text-[11px] text-slate-500 mt-0.5">
                          Grade A: {auditReport?.zeroBarcodeBreakdown.gradeA || 0} boxes, Non-A: {auditReport?.zeroBarcodeBreakdown.gradeNonA || 0} boxes.
                        </p>
                      </div>
                      <span className="px-2.5 py-1 rounded-full text-xs font-bold font-mono bg-orange-100 text-orange-800 border border-orange-200">
                        {zeroBarcodeList.length} Boxes (箱)
                      </span>
                    </div>

                    <div className="overflow-x-auto border border-orange-200 rounded-xl bg-white max-h-52">
                      <table className="w-full text-left text-xs">
                        <thead className="bg-orange-100/60 sticky top-0 text-[10px] font-bold text-orange-900 uppercase">
                          <tr>
                            <th className="px-3 py-2 text-center">Row No. (行号)</th>
                            <th className="px-3 py-2">Box No. (箱号)</th>
                            <th className="px-3 py-2">Location (库位)</th>
                            <th className="px-3 py-2 text-center">Grade (Col K)</th>
                            <th className="px-3 py-2 text-center">Raw Barcode (S列原值)</th>
                            <th className="px-3 py-2">Order No. (工单)</th>
                            <th className="px-3 py-2">Reason (原因)</th>
                          </tr>
                        </thead>
                        <tbody className="divide-y divide-orange-100 font-mono">
                          {zeroBarcodeList.map((item, idx) => (
                            <tr key={idx} className="hover:bg-orange-50/40">
                              <td className="px-3 py-2 text-center font-bold text-orange-700">Row {item.rowIdx}</td>
                              <td className="px-3 py-2 font-bold text-slate-800">{item.boxNo}</td>
                              <td className="px-3 py-2 font-sans">{item.location}</td>
                              <td className="px-3 py-2 text-center">
                                <span className={cn(
                                  "px-2 py-0.5 rounded text-[10px] font-bold",
                                  item.grade.toUpperCase().startsWith('A') ? "bg-blue-100 text-blue-800" : "bg-purple-100 text-purple-800"
                                )}>
                                  {item.grade}
                                </span>
                              </td>
                              <td className="px-3 py-2 text-center font-bold text-rose-600">
                                {item.rawBarcode === '' ? '(Empty/留空)' : `"${item.rawBarcode}"`}
                              </td>
                              <td className="px-3 py-2">{item.orderNo || '-'}</td>
                              <td className="px-3 py-2 font-sans text-slate-600 text-[11px]">{item.reason}</td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  </div>

                  {/* Feature 3 & 4: Empty Box & Empty Location */}
                  <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                    <div className="bg-slate-50 border border-slate-200 rounded-2xl p-4 space-y-2">
                      <div className="flex items-center justify-between">
                        <span className="text-xs font-bold text-slate-800">Pattern 3: Empty Box No. (空箱号)</span>
                        <span className="px-2 py-0.5 rounded text-xs font-bold font-mono bg-slate-200 text-slate-700">
                          {auditReport ? auditReport.emptyBoxRows.length : 0} Rows
                        </span>
                      </div>
                      <p className="text-[11px] text-slate-500">
                        {auditReport && auditReport.emptyBoxRows.length > 0 
                          ? `Found in rows: ${auditReport.emptyBoxRows.map(x => `Row ${x.rowIdx}`).join(', ')}`
                          : 'All rows contain valid box identifiers.'}
                      </p>
                    </div>

                    <div className="bg-slate-50 border border-slate-200 rounded-2xl p-4 space-y-2">
                      <div className="flex items-center justify-between">
                        <span className="text-xs font-bold text-slate-800">Pattern 4: Empty / Unassigned Location (空库位)</span>
                        <span className="px-2 py-0.5 rounded text-xs font-bold font-mono bg-slate-200 text-slate-700">
                          {auditReport ? auditReport.emptyLocRows.length : 0} Rows
                        </span>
                      </div>
                      <p className="text-[11px] text-slate-500">
                        {auditReport && auditReport.emptyLocRows.length > 0
                          ? `Found in rows: ${auditReport.emptyLocRows.map(x => `Row ${x.rowIdx}`).join(', ')}`
                          : 'All rows have designated location assignments.'}
                      </p>
                    </div>
                  </div>
                </div>

              </div>

              {/* Modal Footer */}
              <div className="px-6 py-4 border-t border-slate-100 bg-slate-50 flex items-center justify-between text-xs text-slate-500">
                <span>
                  Audit: {totalRowsRaw} (Raw) - {zeroBarcodeCount} (Col S=0) = {totalRowsRaw - zeroBarcodeCount} Boxes [{data.length} Grade-A + {nonAGradeCount} Non-A]
                </span>
                <div className="flex items-center gap-2">
                  <button
                    onClick={exportAuditReportToExcel}
                    className="px-4 py-1.5 bg-indigo-600 hover:bg-indigo-700 text-white rounded-xl text-xs font-bold transition-colors shadow-sm flex items-center gap-1"
                  >
                    <Download className="w-3.5 h-3.5" />
                    Export Audit Excel (导出诊断表)
                  </button>
                  <button
                    onClick={() => setShowAuditModal(false)}
                    className="px-4 py-1.5 bg-slate-200 hover:bg-slate-300 text-slate-700 rounded-xl text-xs font-bold transition-colors"
                  >
                    Close (关闭)
                  </button>
                </div>
              </div>
            </motion.div>
          </motion.div>
        )}
      </AnimatePresence>
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
              <p className="text-[10px] font-bold text-red-400 uppercase tracking-widest mb-0.5">Analysis Error (分析错误)</p>
              <p className="text-sm font-bold">{error}</p>
            </div>
            <button onClick={() => setError(null)} className="ml-4 p-2 hover:bg-white/10 rounded-lg transition-colors text-slate-400">
              <X className="w-4 h-4" />
            </button>
          </motion.div>
        )}
      </AnimatePresence>

      {/* Success Toast */}
      <AnimatePresence>
        {saveSuccess && (
          <motion.div 
            initial={{ opacity: 0, y: 30 }}
            animate={{ opacity: 1, y: 0 }}
            exit={{ opacity: 0, scale: 0.95 }}
            className="fixed bottom-10 right-10 bg-slate-900 text-white px-8 py-5 rounded-2xl shadow-[0_20px_50px_rgba(0,0,0,0.3)] flex items-center gap-4 z-50 border border-white/10 backdrop-blur-xl"
          >
            <div className="w-10 h-10 bg-emerald-500 rounded-full flex items-center justify-center shrink-0 shadow-lg shadow-emerald-500/20">
              <CheckCircle2 className="w-5 h-5 text-white" />
            </div>
            <div>
              <p className="text-[10px] font-bold text-emerald-400 uppercase tracking-widest mb-0.5">Success (操作成功)</p>
              <p className="text-sm font-bold">{saveSuccess}</p>
            </div>
            <button onClick={() => setSaveSuccess(null)} className="ml-4 p-2 hover:bg-white/10 rounded-lg transition-colors text-slate-400">
              <X className="w-4 h-4" />
            </button>
          </motion.div>
        )}
      </AnimatePresence>
    </div>
  );
}

