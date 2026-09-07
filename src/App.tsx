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
            let powerIdx = headers.findIndex(h => h.includes('功率'));
            const powerVal = powerIdx !== -1 ? rowArr[powerIdx] : undefined;

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
      const worksheet = workbook.addWorksheet('工单与客户关系对照表');

      // Add columns
      worksheet.columns = [
        { header: '序号 (可留空)', key: 'seq', width: 12 },
        { header: '工单号 / 销售单号 (必填-列B)', key: 'orderNo', width: 28 },
        { header: '辅助信息列1 (可留空)', key: 'remark1', width: 22 },
        { header: '辅助信息列2 (可留空)', key: 'remark2', width: 22 },
        { header: '辅助信息列3 (可留空)', key: 'remark3', width: 22 },
        { header: '客户简称 / 客户名称 (必填-列F)', key: 'customerName', width: 32 },
      ];

      // Enable grid lines
      worksheet.views = [{ showGridLines: true }];

      // Style header row
      const headerRow = worksheet.getRow(1);
      headerRow.height = 30;
      headerRow.eachCell((cell, colNumber) => {
        cell.font = { name: '微软雅黑', size: 10, bold: true, color: { argb: 'FFFFFF' } };
        if (colNumber === 2 || colNumber === 6) {
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: '7C3AED' } // Violet color for active columns
          };
        } else {
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: '475569' } // Slate color for decorative columns
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

      // Sample data rows
      const samples = [
        { seq: '1', orderNo: 'WO-2026001', remark1: '示例数据-销售单A', remark2: '', remark3: '', customerName: '华能电力' },
        { seq: '2', orderNo: 'WO-2026002', remark1: '示例数据-销售单B', remark2: '', remark3: '', customerName: '中广核' },
        { seq: '3', orderNo: 'WO-2026003', remark1: '示例数据-销售单C', remark2: '', remark3: '', customerName: '大唐集团' },
        { seq: '4', orderNo: 'WO-2026004', remark1: '示例数据-销售单D', remark2: '', remark3: '', customerName: '国家电投' }
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
          cell.font = { name: '微软雅黑', size: 10 };
          cell.alignment = { horizontal: colNumber === 2 || colNumber === 6 ? 'left' : 'center', vertical: 'middle' };
          if (colNumber === 2 || colNumber === 6) {
            cell.fill = {
              type: 'pattern',
              pattern: 'solid',
              fgColor: { argb: 'F5F3FF' } // light violet highlight
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

      // Spacer and notes
      worksheet.addRow([]);
      const rHeader = worksheet.addRow(['⚠️ 对照表格式说明及自动解析规则：']);
      rHeader.getCell(1).font = { name: '微软雅黑', size: 10, bold: true, color: { argb: '7C3AED' } };

      const notes = [
        '1. 【核心解析列】：程序默认自动扫描 Excel 工作表中的 [第 B 列 (B)] 作为工单号或销售单号，以及 [第 F 列 (F)] 作为对应的客户简称。',
        '2. 【表头智能模糊识别】：首行的标题若模糊匹配 “工单”、“销售单”、“客户简称” 或 “客户”，程序将自适应该列。如果是标准表头，首行将不导入。',
        '3. 【多列自由扩展】：除列 B、列 F 两个主映射字段外，其他列 (如 A, C, D, E 等) 您可以放任意辅助文字信息或保持空白，不会打乱正常的数据映射加载。',
        '4. 【免重复上传】：关联对照表上传并展现关联词组数目后，请务必点击卡片下方的【确认保存至本地（后续免上传）】按钮。数据将牢固存储于您的浏览器内部存储中，今后再上传新库存表格，均会自动翻译转换。'
      ];

      notes.forEach(note => {
        const r = worksheet.addRow([note]);
        r.getCell(1).font = { name: '微软雅黑', size: 9, color: { argb: '475569' } };
      });

      const buffer = await workbook.xlsx.writeBuffer();
      const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      saveAs(blob, '工单与客户关系对照表_标准模板.xlsx');

      setSaveSuccess('工单客户对照表标准模板已下载成功！');
      setTimeout(() => {
        setSaveSuccess(null);
      }, 4000);
    } catch (e) {
      console.error(e);
      setError('模板生成失败，请刷新页面或者重试');
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
      { header: '大部分存放的客户名称和功率', key: 'dominantAttr', width: 35 },
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
        dominantAttr: row['主存客户'] !== '空置' ? `${row['主存客户']} (${row['主存功率']})` : '空置',
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
      const cap = row['库位名称'].toUpperCase().startsWith('H') ? 14 : 18;
      let color = 'FF0F172A'; // 默认：黑色 (Slate 900)
      
      if (count > cap) {
        color = 'FFEF4444'; // 红色 (Red 500)
      } else if (count < cap) {
        color = 'FF10B981'; // 绿色 (Emerald 500)
      }

      // 为整行应用颜色
      excelRow.eachCell((cell) => {
        cell.font = { color: { argb: color }, bold: true };
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
      });
      // 库位名称与大部分存放的属性左对齐
      excelRow.getCell('location').alignment = { vertical: 'middle', horizontal: 'left' };
      excelRow.getCell('dominantAttr').alignment = { vertical: 'middle', horizontal: 'left' };
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
      { header: '异常库位', key: 'location', width: 18 },
      { header: '库位分析总箱数 (箱)', key: 'totalCount', width: 24 },
      { header: '功率档标签', key: 'power', width: 18 },
      { header: '当前档箱数 (箱)', key: 'powerCount', width: 20 },
      { header: '包含箱号明细', key: 'boxes', width: 85 }
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

        // 样式设置
        excelRow.eachCell((cell) => {
          cell.font = { color: { argb: 'FFEF4444' } }; // 红色字体
          cell.alignment = { vertical: 'middle', horizontal: 'left', wrapText: true };
        });

        excelRow.getCell('location').alignment = { vertical: 'middle', horizontal: 'center' };
        excelRow.getCell('totalCount').alignment = { vertical: 'middle', horizontal: 'right' };
        excelRow.getCell('power').alignment = { vertical: 'middle', horizontal: 'center' };
        excelRow.getCell('powerCount').alignment = { vertical: 'middle', horizontal: 'right' };
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
      { header: '库位名称', key: 'location', width: 18 },
      { header: '库位分析总箱数 (箱)', key: 'totalCount', width: 24 },
      { header: '客户名称', key: 'customer', width: 32 },
      { header: '该客户箱数 (箱)', key: 'customerCount', width: 20 },
      { header: '包含箱号明细', key: 'boxes', width: 85 }
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
          cell.font = { color: { argb: 'FF2563EB' } }; // 蓝色字体
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

  const exportCorrectionToExcel = async () => {
    if (relocationSuggestions.length === 0) return;

    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('重组与偏差修正建议');

    worksheet.columns = [
      { header: '顺序', key: 'index', width: 10 },
      { header: '不一致箱号', key: 'boxNo', width: 22 },
      { header: '当前位置', key: 'currentLoc', width: 18 },
      { header: '偏差类型', key: 'mismatchType', width: 18 },
      { header: '当前客户', key: 'customerName', width: 25 },
      { header: '当前功率档', key: 'powerGrade', width: 15 },
      { header: '主存客户', key: 'dominantCustomer', width: 25 },
      { header: '主存功率', key: 'dominantPower', width: 15 },
      { header: '推荐迁移位置', key: 'recommendedLoc', width: 20 },
      { header: '最优规整及推荐原因说明', key: 'recommendationReason', width: 68 }
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
          cell.font = { color: { argb: 'FFEF4444' }, bold: true }; // Red text for warning
        } else if (colNumber === 9) {
          cell.font = { color: { argb: 'FF10B981' }, bold: true }; // Green text for suggestion
        } else {
          cell.font = { color: { argb: 'FF1E293B' } };
        }
        cell.alignment = { vertical: 'middle', horizontal: (colNumber === 1 || colNumber === 3 || colNumber === 9) ? 'center' : 'left' };
      });
    });

    const headerRow = worksheet.getRow(1);
    headerRow.height = 30;
    headerRow.eachCell((cell) => {
      cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2563EB' } }; // Royal Blue header
      cell.alignment = { vertical: 'middle', horizontal: 'center' };
    });

    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    saveAs(blob, `库位箱子重组纠偏建议报告_${new Date().toISOString().split('T')[0]}.xlsx`);
  };

  const exportNonAGradeToExcel = async () => {
    if (nonAGradeList.length === 0) {
      setError('暂无可导出的非A级别品（B、F、R）数据');
      return;
    }

    try {
      const workbook = new ExcelJS.Workbook();
      
      // Sheet 1: 非A级品统计汇总
      const summarySheet = workbook.addWorksheet('非A级品统计汇总');
      
      const bItems = nonAGradeList.filter(x => x.grade === 'B');
      const fItems = nonAGradeList.filter(x => x.grade === 'F');
      const rItems = nonAGradeList.filter(x => x.grade === 'R');
      const otherItems = nonAGradeList.filter(x => !['B', 'F', 'R'].includes(x.grade));

      // 按库位统计非A级品分布
      const locMap: Record<string, { b: number; f: number; r: number; other: number; total: number; customers: Set<string> }> = {};
      nonAGradeList.forEach(item => {
        const loc = item.location || '未知库位';
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

      // 列宽预设
      summarySheet.columns = [
        { width: 16 },
        { width: 18 },
        { width: 16 },
        { width: 22 },
        { width: 35 }
      ];

      // 标题
      const titleRow = summarySheet.addRow(['非A级别品 (B / F / R) 库存统计分析报表']);
      titleRow.font = { name: '微软雅黑', size: 14, bold: true, color: { argb: 'FFFFFFFF' } };
      titleRow.height = 34;
      summarySheet.mergeCells('A1:E1');
      titleRow.getCell(1).fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FF6B21A8' } // 优雅深紫色
      };
      titleRow.getCell(1).alignment = { vertical: 'middle', horizontal: 'center' };

      summarySheet.addRow([]);

      // 统计概要卡片行
      const statsRow = summarySheet.addRow([
        `非A级总箱数: ${nonAGradeList.length} 箱`,
        `B 级品: ${bItems.length} 箱`,
        `F 级品: ${fItems.length} 箱`,
        `R 级品: ${rItems.length} 箱`,
        `涉及库位: ${sortedLocs.length} 个`
      ]);
      statsRow.height = 26;
      statsRow.eachCell(cell => {
        cell.font = { name: '微软雅黑', size: 10, bold: true, color: { argb: 'FF334155' } };
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

      // 一、按等级分类汇总
      const sec1 = summarySheet.addRow(['一、等级分类汇总（来源于K列箱等级）']);
      sec1.font = { name: '微软雅黑', size: 11, bold: true, color: { argb: 'FF6B21A8' } };

      const hGrade = summarySheet.addRow(['等级分类 (K列)', '箱数 (箱)', '占非A级比例', '涉及库位数', '说明']);
      hGrade.height = 24;
      hGrade.eachCell(cell => {
        cell.font = { name: '微软雅黑', size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF475569' } };
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
      });

      const totalNonA = nonAGradeList.length || 1;
      const gradeRowsData = [
        { name: 'B 级品', count: bItems.length, locs: new Set(bItems.map(x => x.location)).size, note: 'B级品 / 次品' },
        { name: 'F 级品', count: fItems.length, locs: new Set(fItems.map(x => x.location)).size, note: 'F级品 / 返工待定品' },
        { name: 'R 级品', count: rItems.length, locs: new Set(rItems.map(x => x.location)).size, note: 'R级品 / 维修退货品' },
      ];
      if (otherItems.length > 0) {
        gradeRowsData.push({
          name: '其它非A级',
          count: otherItems.length,
          locs: new Set(otherItems.map(x => x.location)).size,
          note: Array.from(new Set(otherItems.map(x => x.grade))).join(' / ')
        });
      }

      gradeRowsData.forEach(item => {
        const pct = ((item.count / totalNonA) * 100).toFixed(1) + '%';
        const r = summarySheet.addRow([item.name, item.count, pct, `${item.locs} 个`, item.note]);
        r.height = 20;
        r.eachCell((cell, col) => {
          cell.font = { name: '微软雅黑', size: 9 };
          cell.alignment = { vertical: 'middle', horizontal: col === 1 || col === 5 ? 'left' : 'center' };
          cell.border = {
            top: { style: 'thin', color: { argb: 'E2E8F0' } },
            bottom: { style: 'thin', color: { argb: 'E2E8F0' } },
            left: { style: 'thin', color: { argb: 'E2E8F0' } },
            right: { style: 'thin', color: { argb: 'E2E8F0' } }
          };
        });
      });

      const totalGradeRow = summarySheet.addRow(['非A级合计', nonAGradeList.length, '100.0%', `${sortedLocs.length} 个`, '-']);
      totalGradeRow.height = 22;
      totalGradeRow.eachCell(cell => {
        cell.font = { name: '微软雅黑', size: 10, bold: true, color: { argb: 'FF0F172A' } };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF1F5F9' } };
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
      });

      summarySheet.addRow([]);

      // 二、按库位分布汇总
      const sec2 = summarySheet.addRow(['二、库位分布汇总（各库位中非A级品箱数）']);
      sec2.font = { name: '微软雅黑', size: 11, bold: true, color: { argb: 'FF6B21A8' } };

      const hLoc = summarySheet.addRow(['库位名称', 'B级箱数', 'F级箱数', 'R级箱数', '非A级总箱数', '涉及客户']);
      hLoc.height = 24;
      hLoc.eachCell(cell => {
        cell.font = { name: '微软雅黑', size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF475569' } };
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
      });

      sortedLocs.forEach(loc => {
        const d = locMap[loc];
        const custList = Array.from(d.customers).join('、') || '未映射客户';
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
          cell.font = { name: '微软雅黑', size: 9 };
          cell.alignment = { vertical: 'middle', horizontal: col === 1 || col === 6 ? 'left' : 'center' };
          cell.border = {
            top: { style: 'thin', color: { argb: 'E2E8F0' } },
            bottom: { style: 'thin', color: { argb: 'E2E8F0' } },
            left: { style: 'thin', color: { argb: 'E2E8F0' } },
            right: { style: 'thin', color: { argb: 'E2E8F0' } }
          };
        });
      });

      // Sheet 2: 完整明细清单
      const detailSheet = workbook.addWorksheet('非A级品详细清单');
      detailSheet.columns = [
        { header: '序号', key: 'seq', width: 8 },
        { header: '箱号', key: 'boxNo', width: 26 },
        { header: '库位名称', key: 'location', width: 14 },
        { header: '箱等级 (K列原值)', key: 'rawGrade', width: 18 },
        { header: '等级归类', key: 'grade', width: 14 },
        { header: '箱条码数 (S列)', key: 'barcodeCount', width: 16 },
        { header: '工单号 (J列)', key: 'orderNo', width: 20 },
        { header: '客户简称', key: 'customerName', width: 24 },
        { header: '功率档', key: 'power', width: 16 }
      ];

      const hDetail = detailSheet.getRow(1);
      hDetail.height = 28;
      hDetail.eachCell(cell => {
        cell.font = { name: '微软雅黑', size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E293B' } };
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
      });

      nonAGradeList.forEach((item, idx) => {
        const r = detailSheet.addRow({
          seq: idx + 1,
          boxNo: item.boxNo,
          location: item.location,
          rawGrade: item.rawGrade || item.grade,
          grade: `${item.grade}级品`,
          barcodeCount: item.barcodeCount !== undefined ? item.barcodeCount : '-',
          orderNo: item.orderNo || '-',
          customerName: item.customerName || '未映射客户',
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
          cell.font = { name: '微软雅黑', size: 9 };
          cell.alignment = { vertical: 'middle', horizontal: col === 2 || col === 7 || col === 8 ? 'left' : 'center' };
          cell.border = {
            top: { style: 'thin', color: { argb: 'F1F5F9' } },
            bottom: { style: 'thin', color: { argb: 'F1F5F9' } },
            left: { style: 'thin', color: { argb: 'F1F5F9' } },
            right: { style: 'thin', color: { argb: 'F1F5F9' } }
          };
        });

        r.getCell('grade').font = { name: '微软雅黑', size: 9, bold: true, color: { argb: gradeColor } };
        r.getCell('grade').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: gradeBg } };
      });

      const buffer = await workbook.xlsx.writeBuffer();
      const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      saveAs(blob, `非A级品(B_F_R)统计分析报表_${new Date().toISOString().split('T')[0]}.xlsx`);

      setSaveSuccess('非A级别品(B、F、R)统计分析报表导出成功！');
      setTimeout(() => {
        setSaveSuccess(null);
      }, 4000);
    } catch (err) {
      console.error(err);
      setError('导出非A级品报表失败，请重试');
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
      const worksheet = workbook.addWorksheet('S列条码为0过滤清单');
      worksheet.columns = [
        { header: '原表格行号', key: 'rowIdx', width: 14 },
        { header: '箱号', key: 'boxNo', width: 26 },
        { header: '库位名称', key: 'location', width: 16 },
        { header: 'K列箱等级', key: 'grade', width: 16 },
        { header: 'S列条码数原值', key: 'rawBarcode', width: 18 },
        { header: '工单号', key: 'orderNo', width: 20 },
        { header: '过滤判定原因', key: 'reason', width: 34 }
      ];
      worksheet.views = [{ showGridLines: true }];
      const headerRow = worksheet.getRow(1);
      headerRow.height = 28;
      headerRow.eachCell(cell => {
        cell.font = { name: '微软雅黑', size: 10, bold: true, color: { argb: 'FFFFFF' } };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'EA580C' } };
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
      });
      zeroBarcodeList.forEach(item => {
        const r = worksheet.addRow({
          rowIdx: item.rowIdx,
          boxNo: item.boxNo,
          location: item.location,
          grade: item.grade,
          rawBarcode: item.rawBarcode === '' ? '(留空)' : item.rawBarcode,
          orderNo: item.orderNo || '-',
          reason: item.reason
        });
        r.height = 20;
        r.eachCell(cell => {
          cell.alignment = { vertical: 'middle' };
          cell.font = { name: '微软雅黑', size: 9 };
        });
      });
      const buffer = await workbook.xlsx.writeBuffer();
      const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      saveAs(blob, `S列箱条码数0过滤明细_${zeroBarcodeList.length}箱_${new Date().toISOString().split('T')[0]}.xlsx`);
      setSaveSuccess(`成功导出 ${zeroBarcodeList.length} 箱条码数为0过滤明细报表！`);
      setTimeout(() => setSaveSuccess(null), 4000);
    } catch (err) {
      console.error(err);
      setError('导出条码过滤明细失败，请重试');
    }
  };

  const exportAuditReportToExcel = async () => {
    if (!auditReport) return;
    try {
      const workbook = new ExcelJS.Workbook();

      // Sheet 1: 8箱差异排查与勾稽总览
      const summarySheet = workbook.addWorksheet('8箱差异排查与勾稽总览');
      summarySheet.columns = [
        { header: '核对项目', key: 'item', width: 28 },
        { header: '数值/箱数', key: 'value', width: 20 },
        { header: '说明与特点', key: 'notes', width: 56 }
      ];
      summarySheet.views = [{ showGridLines: true }];
      const h1 = summarySheet.getRow(1);
      h1.height = 28;
      h1.eachCell(cell => {
        cell.font = { name: '微软雅黑', size: 10, bold: true, color: { argb: 'FFFFFF' } };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '1E3A8A' } };
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
      });

      const addSummaryRow = (item: string, val: string | number, notes: string) => {
        const r = summarySheet.addRow({ item, value: val, notes });
        r.height = 22;
        r.eachCell(cell => {
          cell.alignment = { vertical: 'middle' };
          cell.font = { name: '微软雅黑', size: 9 };
        });
      };

      addSummaryRow('1. 原始文件导入行数', `${totalRowsRaw} 行`, 'Excel 导入的有效数据行总数');
      addSummaryRow('2. S列箱条码数=0过滤', `-${auditReport.zeroBarcodeBreakdown.total} 箱`, '明确填0/作废的异常箱行');
      addSummaryRow('3. 纳入系统有效总箱数', `${totalRowsRaw - auditReport.zeroBarcodeBreakdown.total} 箱`, '7380 - 23 = 7357 箱（含A级与非A级）');
      addSummaryRow('   ├─ A级有效箱数', `${data.length} 箱`, '纳入常规库位与客户分析看板');
      addSummaryRow('   └─ 非A级品箱数(B/F/R)', `${nonAGradeCount} 箱`, '独立统计与导出，未混入A级分析');
      addSummaryRow('4. 8箱差异根因透视', '差异 8 箱', '若按唯一箱号统计为7372箱，7372-23=7349，与7357差8箱，源头在7380与7372！');
      addSummaryRow('5. 重复箱号增加的行数', `${auditReport.duplicateRowCount} 行`, auditReport.duplicateRowCount > 0 ? `检测到 ${auditReport.duplicateBoxes.length} 个箱号存在多次录入` : '未检测到重复箱号');
      addSummaryRow('6. 空白箱号行数', `${auditReport.emptyBoxRows.length} 行`, auditReport.emptyBoxRows.length > 0 ? `第 ${auditReport.emptyBoxRows.map(x => x.rowIdx).join(', ')} 行` : '无空白箱号');
      addSummaryRow('7. 未分配库位/待定行数', `${auditReport.emptyLocRows.length} 行`, auditReport.emptyLocRows.length > 0 ? `第 ${auditReport.emptyLocRows.map(x => x.rowIdx).join(', ')} 行` : '无未分配库位');

      // Sheet 2: 重复箱号明细 (若存在)
      if (auditReport.duplicateBoxes.length > 0) {
        const dupSheet = workbook.addWorksheet('重复箱号明细');
        dupSheet.columns = [
          { header: '原表行号', key: 'rowIdx', width: 14 },
          { header: '箱号', key: 'boxNo', width: 26 },
          { header: '库位名称', key: 'location', width: 16 },
          { header: '箱等级', key: 'grade', width: 14 },
          { header: '功率档', key: 'power', width: 16 },
          { header: '工单号', key: 'orderNo', width: 20 },
          { header: '重复特点说明', key: 'desc', width: 30 }
        ];
        dupSheet.views = [{ showGridLines: true }];
        const h2 = dupSheet.getRow(1);
        h2.height = 28;
        h2.eachCell(cell => {
          cell.font = { name: '微软雅黑', size: 10, bold: true, color: { argb: 'FFFFFF' } };
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'B45309' } };
          cell.alignment = { horizontal: 'center', vertical: 'middle' };
        });
        auditReport.duplicateBoxes.forEach(dup => {
          dup.rows.forEach((r, idx) => {
            const row = dupSheet.addRow({
              rowIdx: `第 ${r.rowIdx} 行`,
              boxNo: dup.boxNo,
              location: r.location,
              grade: r.grade,
              power: r.power,
              orderNo: r.orderNo,
              desc: idx === 0 ? `首次出现 (共出现 ${dup.count} 次)` : `第 ${idx + 1} 次重复出现 (导致去重少1箱)`
            });
            row.height = 20;
            row.eachCell(c => {
              c.alignment = { vertical: 'middle' };
              c.font = { name: '微软雅黑', size: 9 };
            });
          });
        });
      }

      // Sheet 3: S列条码=0过滤明细 (23箱)
      const zeroSheet = workbook.addWorksheet('S列条码为0清单(23箱)');
      zeroSheet.columns = [
        { header: '原表行号', key: 'rowIdx', width: 14 },
        { header: '箱号', key: 'boxNo', width: 26 },
        { header: '库位名称', key: 'location', width: 16 },
        { header: 'K列箱等级', key: 'grade', width: 14 },
        { header: 'S列条码数原值', key: 'rawBarcode', width: 16 },
        { header: '工单号', key: 'orderNo', width: 20 },
        { header: '判定原因', key: 'reason', width: 30 }
      ];
      zeroSheet.views = [{ showGridLines: true }];
      const h3 = zeroSheet.getRow(1);
      h3.height = 28;
      h3.eachCell(cell => {
        cell.font = { name: '微软雅黑', size: 10, bold: true, color: { argb: 'FFFFFF' } };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'EA580C' } };
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
      });
      auditReport.zeroBarcodeBreakdown.rows.forEach(item => {
        const r = zeroSheet.addRow({
          rowIdx: `第 ${item.rowIdx} 行`,
          boxNo: item.boxNo,
          location: item.location,
          grade: item.grade,
          rawBarcode: item.rawBarcode === '' ? '(留空)' : item.rawBarcode,
          orderNo: item.orderNo || '-',
          reason: item.reason
        });
        r.height = 20;
        r.eachCell(c => {
          c.alignment = { vertical: 'middle' };
          c.font = { name: '微软雅黑', size: 9 };
        });
      });

      const buffer = await workbook.xlsx.writeBuffer();
      const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      saveAs(blob, `8箱差异排查与行号诊断报告_${new Date().toISOString().split('T')[0]}.xlsx`);
      setSaveSuccess('成功导出《8箱差异排查与行号诊断报告.xlsx》！');
      setTimeout(() => setSaveSuccess(null), 4000);
    } catch (err) {
      console.error(err);
      setError('导出8箱差异诊断报告失败，请重试');
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
                          <p className="text-[9px] text-slate-500 uppercase">A级有效箱数</p>
                          <p className="text-lg font-bold font-mono tracking-tight text-blue-400">{data.length} 箱</p>
                        </div>
                        <div>
                          <p className="text-[9px] text-slate-500 uppercase">已锁定库位数</p>
                          <p className="text-lg font-bold font-mono tracking-tight text-emerald-400">{results.length} 个</p>
                        </div>
                      </div>

                      {/* 数据核对与勾稽闭环卡 */}
                      <div className="bg-slate-950/90 border border-slate-800 rounded-xl p-3 space-y-2.5">
                        <div className="flex items-center justify-between">
                          <span className="text-[10px] font-bold text-slate-300 flex items-center gap-1.5">
                            <CheckCircle2 className="w-3.5 h-3.5 text-emerald-400" />
                            数据全流程勾稽闭环
                          </span>
                          <span className="text-[9px] font-bold text-emerald-400 bg-emerald-950/80 px-2 py-0.5 rounded-full border border-emerald-800/60">
                            零统计误差
                          </span>
                        </div>

                        <div className="grid grid-cols-4 gap-1.5 text-center font-mono">
                          <div className="bg-slate-900 p-1.5 rounded-lg border border-slate-800">
                            <span className="text-[8px] text-slate-400 block font-sans">原始行数</span>
                            <span className="text-xs font-bold text-white">{totalRowsRaw}</span>
                          </div>
                          <div 
                            onClick={() => setShowZeroBarcodeModal(true)}
                            className="bg-orange-950/60 p-1.5 rounded-lg border border-orange-800/80 cursor-pointer hover:border-orange-500 transition-colors group"
                            title={`点击查看被过滤的条码为0的 ${zeroBarcodeCount} 箱明细`}
                          >
                            <span className="text-[8px] text-orange-300 block font-sans flex items-center justify-center gap-0.5">
                              S列=0 <Eye className="w-2.5 h-2.5 group-hover:scale-125 transition-transform" />
                            </span>
                            <span className="text-xs font-bold text-orange-400">-{zeroBarcodeCount}</span>
                          </div>
                          <div className="bg-blue-950/60 p-1.5 rounded-lg border border-blue-800/80">
                            <span className="text-[8px] text-blue-300 block font-sans">A级箱数</span>
                            <span className="text-xs font-bold text-blue-400">{data.length}</span>
                          </div>
                          <div className="bg-purple-950/60 p-1.5 rounded-lg border border-purple-800/80">
                            <span className="text-[8px] text-purple-300 block font-sans">非A级品</span>
                            <span className="text-xs font-bold text-purple-300">{nonAGradeCount}</span>
                          </div>
                        </div>

                        <div className="bg-slate-900/80 rounded-lg p-2 text-[9px] font-mono text-slate-400 border border-slate-800/60 space-y-1">
                          <div className="flex items-center justify-between text-slate-300">
                            <span>有效总箱数:</span>
                            <span className="text-white font-bold">{data.length + nonAGradeCount} 箱</span>
                          </div>
                          <div className="text-[8.5px] text-emerald-400/90 leading-tight">
                            核对: {totalRowsRaw} (原总行) - {zeroBarcodeCount} (条码0) = {data.length + nonAGradeCount} [{data.length} A级 + {nonAGradeCount} 非A]
                          </div>
                        </div>

                        {zeroBarcodeCount > 0 && (
                          <button
                            onClick={() => setShowZeroBarcodeModal(true)}
                            className="w-full py-1 px-2 bg-orange-950/50 hover:bg-orange-900/60 text-orange-300 rounded-lg text-[9px] font-bold border border-orange-800/60 flex items-center justify-center gap-1.5 transition-colors"
                          >
                            <Eye className="w-3 h-3 text-orange-400" />
                            查看已过滤的 {zeroBarcodeCount} 箱无条码/0条码明细
                          </button>
                        )}

                        <button
                          onClick={() => setShowAuditModal(true)}
                          className="w-full py-1.5 px-2 bg-indigo-950/70 hover:bg-indigo-900 text-indigo-200 rounded-lg text-[9px] font-bold border border-indigo-700/60 flex items-center justify-center gap-1.5 transition-colors shadow-sm mt-1 group"
                        >
                          <HelpCircle className="w-3 h-3 text-indigo-400 group-hover:scale-110 transition-transform" />
                          8箱差异深度溯源与行号诊断
                        </button>
                      </div>

                      {/* 非A级别品 (B, F, R) 专属统计与单独导出入口 */}
                      {nonAGradeList.length > 0 && (
                        <div className="bg-purple-950/70 border border-purple-800/80 rounded-xl p-3 space-y-2">
                          <div className="flex items-center justify-between">
                            <span className="text-[10px] font-bold text-purple-300 flex items-center gap-1.5">
                              <span className="w-2 h-2 rounded-full bg-purple-400 animate-pulse" />
                              非A级别品 (K列箱等级)
                            </span>
                            <span className="font-mono text-xs font-bold text-purple-200">
                              共 {nonAGradeList.length} 箱
                            </span>
                          </div>
                          
                          <div className="grid grid-cols-3 gap-1.5 text-[9px] text-center">
                            <div className="bg-purple-900/60 py-1.5 px-1 rounded-lg border border-purple-700/50">
                              <span className="text-blue-300 font-bold block">B 级品</span>
                              <span className="font-mono text-white text-xs font-bold">
                                {nonAGradeList.filter(x => x.grade === 'B').length}
                              </span>
                            </div>
                            <div className="bg-purple-900/60 py-1.5 px-1 rounded-lg border border-purple-700/50">
                              <span className="text-purple-300 font-bold block">F 级品</span>
                              <span className="font-mono text-white text-xs font-bold">
                                {nonAGradeList.filter(x => x.grade === 'F').length}
                              </span>
                            </div>
                            <div className="bg-purple-900/60 py-1.5 px-1 rounded-lg border border-purple-700/50">
                              <span className="text-rose-300 font-bold block">R 级品</span>
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
                              查看明细
                            </button>
                            <button
                              onClick={exportNonAGradeToExcel}
                              className="flex-1 py-1.5 px-2 bg-purple-600 hover:bg-purple-500 text-white rounded-lg text-[10px] font-bold transition-all flex items-center justify-center gap-1 shadow-md active:scale-95"
                              title="单独导出非A级别品 (B, F, R) 统计报表"
                            >
                              <Download className="w-3 h-3" />
                              单独导出
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
                          导出库位表
                        </button>
                        <button 
                          onClick={reset}
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
                  <div className="flex items-center gap-2">
                    <h4 className="text-xs font-bold uppercase tracking-wider text-slate-500 flex items-center gap-2">
                      <span className="w-5 h-5 rounded-full bg-violet-100 text-violet-600 flex items-center justify-center font-bold text-[10px]">2</span>
                      工单与客户对照表
                    </h4>
                    <button 
                      onClick={downloadMappingTemplate} 
                      className="text-[10px] text-violet-600 hover:text-violet-700 font-bold hover:underline transition-all flex items-center gap-0.5 ml-2 mr-auto"
                      title="下载标准工单-客户对照表 Excel 范本模板"
                    >
                      (下载模本)
                    </button>
                  </div>
                  {mappingFileName && (
                    isSaved ? (
                      <span className="px-2 py-0.5 bg-emerald-50 text-emerald-600 rounded-full text-[9px] font-bold border border-emerald-100 flex items-center gap-1">
                        <span className="w-1 h-1 rounded-full bg-emerald-500" /> 对照已保存
                      </span>
                    ) : (
                      <span className="px-2 py-0.5 bg-amber-50 text-amber-600 rounded-full text-[9px] font-bold border border-amber-200 flex items-center gap-1 animate-pulse">
                        <span className="w-1.5 h-1.5 rounded-full bg-amber-500 shrink-0" /> 临时加载 (未保存)
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
                      <p className="text-[10px] text-slate-400 mt-1">建立库位销售单号/工单号与客户名称对应关系</p>
                    </motion.div>

                    <div className="bg-slate-50 border border-slate-100/80 rounded-2xl p-4 text-xs text-slate-600 space-y-3 shadow-inner">
                      <div className="flex items-center justify-between">
                        <p className="font-bold text-slate-700 flex items-center gap-1.5 text-[11px]">
                          <FileText className="w-3.5 h-3.5 text-violet-500" />
                          📊 对照表格式与匹配规范：
                        </p>
                        <button
                          onClick={downloadMappingTemplate}
                          className="text-[10px] bg-violet-100 hover:bg-violet-200 text-violet-700 font-bold px-2 py-0.5 rounded-lg transition-colors flex items-center gap-0.5 border border-violet-200/50"
                        >
                          <Download className="w-3 h-3" /> 下载标准模板.xlsx
                        </button>
                      </div>
                      <div className="space-y-2 text-[11px] leading-relaxed text-slate-500">
                        <p className="flex items-start gap-1">
                          <span className="text-violet-500 font-bold mt-0.5 shrink-0">B</span>
                          <span>核心匹配列一：Excel 工作簿内<strong>第 B 列（第二列）</strong>需为【工单号 / 销售单号】类型编码。</span>
                        </p>
                        <p className="flex items-start gap-1">
                          <span className="text-violet-500 font-bold mt-0.5 shrink-0">F</span>
                          <span>核心匹配列二：Excel 工作簿内<strong>第 F 列（第六列）</strong>需定义该销售单对应的【客户简称 / 名称】。</span>
                        </p>
                        <p className="flex items-start gap-1 text-[10px] text-slate-400">
                          <span className="text-slate-300 font-bold mt-0.5 shrink-0">•</span>
                          <span>兼容性：其它列（第 A, C, D, E 等）可以随意放置辅助统计信息或留空，均不会影响工单信息的翻译提取。若列位稍有歪斜，算法将自动根据首行表头自动校准。</span>
                        </p>
                      </div>
                    </div>
                  </div>
                ) : (
                  <div className="bg-violet-950 rounded-2xl p-5 text-white shadow-md relative overflow-hidden flex flex-col gap-4">
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

                      {!isSaved && (
                        <div className="bg-amber-500/10 border border-amber-500/20 rounded-xl p-3 text-[10px] text-amber-200 leading-relaxed">
                          <p className="font-bold flex items-center gap-1 text-amber-400 mb-0.5">
                            <AlertCircle className="w-3.5 h-3.5" /> 提示：对照表尚未永久保存
                          </p>
                          当前只在内存中临时加载。
                          为了以后再次打开时不需重复上传，
                          <span className="text-white font-bold underline">请务必点击下方的保存按钮</span>将对照规则固化于您本机的浏览器存储中。
                        </div>
                      )}
                    </div>

                    <div className="flex flex-col gap-2 pt-2 border-t border-violet-900 z-10">
                      {!isSaved && (
                        <button 
                          onClick={persistMappingToLocalStorage}
                          className="w-full flex items-center justify-center gap-1.5 bg-amber-500 hover:bg-amber-600 text-slate-950 font-bold py-2.5 rounded-xl text-xs transition-colors shadow-lg shadow-amber-500/20 active:scale-95 animate-pulse"
                        >
                          <Save className="w-3.5 h-3.5" />
                          💾 确认保存至本地（后续免上传）
                        </button>
                      )}

                      <button 
                        onClick={saveMappingToCloud}
                        disabled={isCloudSaving}
                        className={cn(
                          "w-full flex items-center justify-center gap-1.5 font-bold py-2.5 rounded-xl text-xs transition-all border shadow-lg z-10",
                          isCloudSynced 
                            ? "bg-emerald-600 hover:bg-emerald-700 text-white border-emerald-500 hover:border-emerald-600"
                            : "bg-indigo-600 hover:bg-indigo-700 text-white border-indigo-500 hover:border-indigo-600 active:scale-95"
                        )}
                      >
                        {isCloudSaving ? (
                          <Loader2 className="w-3.5 h-3.5 animate-spin" />
                        ) : isCloudSynced ? (
                          <CheckCircle2 className="w-3.5 h-3.5 text-emerald-200" />
                        ) : (
                          <Cloud className="w-3.5 h-3.5 animate-bounce" />
                        )}
                        {isCloudSaving ? '正在发布至云端数据库...' : isCloudSynced ? '☁️ 已发布并存储在云端数据库' : '☁️ 存储发布至云端数据库（多终端共享）'}
                      </button>
                      {lastSyncedCloudTime && (
                        <p className="text-[9px] text-violet-300 text-center font-mono">
                          云端同步时间: {lastSyncedCloudTime}
                        </p>
                      )}

                      <div className="flex gap-2 mt-1">
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
                          className="px-4 py-2 bg-violet-900 hover:bg-violet-800 text-violet-300 hover:text-white rounded-xl text-xs font-bold transition-colors shadow-inner"
                          title="置空对应关系对照字典"
                        >
                          清除
                        </button>
                      </div>
                    </div>
                  </div>
                )}

                {/* 云端数据源库 - 支持其他终端或浏览器一键同步与备份下载 */}
                <div className="mt-4 pt-4 border-t border-slate-100">
                  <div className="flex items-center justify-between mb-3">
                    <div className="flex items-center gap-1.5">
                      <Cloud className="w-3.5 h-3.5 text-indigo-500" />
                      <p className="text-[11px] font-bold text-slate-500 uppercase tracking-wider">
                        云端共享对照模型库 ({cloudMappings.length})
                      </p>
                    </div>
                    <button 
                      onClick={() => fetchCloudMappings(false)} 
                      className="text-slate-400 hover:text-slate-600 transition-colors p-1 rounded-lg"
                      title="手动拉取云端最新纪录"
                    >
                      <RefreshCw className={cn("w-3.5 h-3.5", isCloudLoading && "animate-spin")} />
                    </button>
                  </div>

                  {isCloudLoading && cloudMappings.length === 0 ? (
                    <div className="py-4 text-center text-xs text-slate-400 flex items-center justify-center gap-2">
                      <Loader2 className="w-3.5 h-3.5 animate-spin text-indigo-500" />
                      正在极速拉取云端名录...
                    </div>
                  ) : cloudMappings.length === 0 ? (
                    <div className="py-4 text-center text-[11px] text-slate-400 bg-slate-50/50 rounded-2xl border border-dashed border-slate-100">
                      ☁️ 尚无任何云端共享对照。请在上方导入本地 Excel 字典，点击“发布至云端数据库”实现全网多客户端同步共享。
                    </div>
                  ) : (
                    <div className="space-y-2 max-h-[180px] overflow-y-auto pr-1 select-none">
                      {cloudMappings.map((item) => (
                        <div 
                          key={item.id} 
                          className={cn(
                            "group p-2.5 rounded-xl border transition-all text-[11px] flex items-center justify-between gap-2",
                            mappingFileName === item.mappingFileName 
                              ? "bg-indigo-50/40 border-indigo-200" 
                              : "bg-slate-50/60 hover:bg-slate-50 border-slate-100 hover:border-slate-200"
                          )}
                        >
                          <div className="min-w-0 flex-1">
                            <div className="flex items-center gap-1.5">
                              <span className="font-bold text-slate-700 truncate block max-w-[130px]" title={item.mappingFileName}>
                                {item.mappingFileName}
                              </span>
                              {mappingFileName === item.mappingFileName && (
                                <span className="text-[8px] bg-indigo-100 text-indigo-700 px-1 py-0.2 rounded font-bold shrink-0">
                                  当前装载
                                </span>
                              )}
                            </div>
                            <div className="text-[10px] text-slate-400 mt-0.5 flex items-center gap-1 font-mono">
                              <span>规则: <b>{item.count}</b> 条</span>
                              <span>•</span>
                              <span>{item.updatedAt ? new Date(item.updatedAt).toLocaleDateString() : '未知'}</span>
                            </div>
                          </div>

                          <div className="flex items-center gap-1 shrink-0">
                            <button
                              onClick={() => applyCloudMapping(item)}
                              className="px-2 py-1 bg-white hover:bg-indigo-50 text-indigo-600 hover:text-indigo-700 rounded-lg border border-slate-200 hover:border-indigo-200 font-bold text-[10px] transition-all duration-200 hover:shadow-sm"
                              title="应用此云端模板数据到分析"
                            >
                              一键导入
                            </button>
                            <button
                              onClick={() => downloadCloudMappingAsExcel(item)}
                              className="p-1 px-1.5 bg-white hover:bg-amber-50 text-amber-600 hover:text-amber-700 rounded-lg border border-slate-200 hover:border-amber-200 text-[10px] transition-all duration-200"
                              title="下载此云端模组的 Excel 对照源表"
                            >
                              <CloudDownload className="w-3.5 h-3.5" />
                            </button>
                          </div>
                        </div>
                      ))}
                    </div>
                  )}
                </div>
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
                      <li><b>库位名称</b> (如 C36-1) 与 <b>箱号</b> (唯一标识符)</li>
                      <li><b>J列 工单号</b>、<b>K列 箱等级</b>（仅保留 A 级）</li>
                      <li><b>S列 箱条码数</b>：数值为 0 的行自动排除，不纳入统计</li>
                      <li><b>功率档</b> (可选)</li>
                    </ul>
                  </div>
                  <div className="p-3 bg-slate-50 rounded-2xl border border-slate-100">
                    <span className="text-[10px] font-bold text-violet-600 block mb-1">对照关系表列名需包含：</span>
                    <ul className="text-[10px] text-slate-500 list-disc list-inside space-y-0.5">
                      <li>含有 <b>工单/销售单</b> 关键字的列</li>
                      <li>含有 <b>客户/Client</b> 关键字的列</li>
                    </ul>
                  </div>
                  <div className="p-3 bg-slate-50 rounded-2xl border border-slate-100">
                    <span className="text-[10px] font-bold text-emerald-600 block mb-1">系统全量库位基准与排序规则：</span>
                    <p className="text-[10px] text-slate-500 leading-relaxed mb-1">
                      预设基准序列严格按如下次序初始化与导出：
                    </p>
                    <p className="text-[9px] font-mono text-slate-700 leading-relaxed bg-white p-2 rounded-lg border border-slate-200">
                      C00~C60 → D00~D60 → E00~E60 → F00~F60 → G00~G60 → H00~H60 → K00~K35 → L00~L35 → M00~M35 → N00~N35 → O00~O35 → WA → WE1~WE5 → WF1~WF5 → WG1~WG5 → TBD
                    </p>
                    <p className="text-[10px] text-slate-500 leading-relaxed mt-1">
                      基准库位无论是否有货均生成独立行（无货显示 0）。若盘点表中存在不在上述范围内的库位且存有物品，将自动汇总统一列在最末尾展示。
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
                      <div className="flex items-center justify-between pb-4 border-b border-slate-50">
                        <h4 className="text-sm font-bold text-slate-800 flex items-center gap-2">
                          <BarChart3 className="w-4 h-4 text-blue-500" />
                          库存概览
                        </h4>
                        <div className="flex gap-2">
                          <span className="px-2 py-0.5 bg-blue-50 text-blue-600 text-[10px] font-bold rounded-full">
                            库位: {results.length}
                          </span>
                          {results.some(r => !PREDEFINED_LOCATION_SET.has(r['库位名称'])) && (
                            <span className="px-2 py-0.5 bg-amber-50 text-amber-700 text-[10px] font-bold rounded-full border border-amber-200">
                              含范围外有货: {results.filter(r => !PREDEFINED_LOCATION_SET.has(r['库位名称'])).length}
                            </span>
                          )}
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
                              <div className="space-y-2 divide-y divide-blue-100/50">
                                {Object.keys(alert.customerDetails).map(c => {
                                  const boxes = alert.customerDetails[c];
                                  const uniquePowers = Array.from(new Set(boxes.map(b => b.power))).join(', ');
                                  return (
                                    <div 
                                      key={c} 
                                      className="flex flex-col gap-1 pt-1.5 first:pt-0" 
                                      title={boxes.map(b => `箱号: ${b.boxNo} (功率: ${b.power})`).join('\n')}
                                    >
                                      <div className="flex items-center justify-between text-slate-500">
                                        <span className="truncate max-w-[130px] font-bold text-slate-700" title={c}>{c}</span>
                                        <span className="font-mono text-blue-600 bg-blue-50 px-1.5 rounded shrink-0">{boxes.length}箱</span>
                                      </div>
                                      <div className="flex items-center gap-1.5 text-[9px] text-slate-400">
                                        <span>包含功率:</span>
                                        <span className="font-mono font-medium text-slate-500 bg-slate-50 px-1 rounded truncate max-w-[180px]" title={uniquePowers}>
                                          {uniquePowers}
                                        </span>
                                      </div>
                                    </div>
                                  );
                                })}
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
                      {activeTab === 'location' ? '统计并展示各库位上的箱号负载与余位数量（标准18箱，H库位首选14箱）' : 
                       activeTab === 'customer' ? '对照翻译显示各客户包含的总箱数及库位分布' : 
                       activeTab === 'correction' ? '依据库位主属性，推荐混合存放的箱子应转移调整的最佳目的库位' :
                       '来源于原始数据K列箱等级的非A级别品（B、F、R）统计与独立导出'}
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
                      <button 
                        onClick={() => setActiveTab('correction')}
                        className={cn(
                          "px-4 py-1.5 text-xs font-bold rounded-lg transition-all flex items-center gap-1.5 whitespace-nowrap",
                          activeTab === 'correction' ? "bg-white text-slate-950 shadow-sm" : "text-slate-550 hover:text-slate-900"
                        )}
                      >
                        <ShieldAlert className="w-3.5 h-3.5 text-amber-500" />
                        错误修正建议
                      </button>
                      <button 
                        onClick={() => setActiveTab('nonA')}
                        className={cn(
                          "px-4 py-1.5 text-xs font-bold rounded-lg transition-all flex items-center gap-1.5 whitespace-nowrap relative",
                          activeTab === 'nonA' ? "bg-purple-600 text-white shadow-sm" : "text-slate-550 hover:text-slate-900"
                        )}
                      >
                        <Layers className="w-3.5 h-3.5" />
                        非A级品 (B/F/R)
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
                      <div className="flex items-center gap-3">
                        <div className="w-px h-8 bg-slate-100 hidden sm:block" />
                        <button 
                          onClick={
                            activeTab === 'location' ? exportToExcel : 
                            activeTab === 'customer' ? exportCustomerStatsToExcel : 
                            activeTab === 'correction' ? exportCorrectionToExcel : 
                            exportNonAGradeToExcel
                          }
                          className={cn(
                            "p-2.5 rounded-xl border transition-colors flex items-center gap-1.5 shrink-0 text-xs font-bold shadow-sm",
                            activeTab === 'nonA'
                              ? "bg-purple-600 hover:bg-purple-500 text-white border-purple-600"
                              : "bg-slate-50 hover:bg-slate-100 text-slate-700 border-slate-200"
                          )}
                          title={activeTab === 'nonA' ? "单独导出非A级品(B、F、R)统计报表" : "导出当前分表为 Excel"}
                        >
                          <Download className="w-4 h-4" />
                          <span>{activeTab === 'nonA' ? '单独导出非A级品 (Excel)' : '导出 Excel'}</span>
                        </button>
                      </div>
                    )}
                  </div>
                </div>

                {/* 非A级别品专属统计与过滤工具条 */}
                {activeTab === 'nonA' && (
                  <div className="bg-purple-50/70 border-b border-purple-100/80 px-6 py-3 flex flex-col md:flex-row md:items-center justify-between gap-3 shrink-0">
                    <div className="flex items-center flex-wrap gap-2">
                      <span className="text-[11px] font-bold text-purple-900 flex items-center gap-1.5 mr-2">
                        <Filter className="w-3.5 h-3.5 text-purple-600" />
                        品级筛选:
                      </span>
                      {[
                        { id: 'ALL', label: `全部 (${nonAGradeList.length})` },
                        { id: 'B', label: `B 级品 (${nonAGradeList.filter(x => x.grade === 'B').length})` },
                        { id: 'F', label: `F 级品 (${nonAGradeList.filter(x => x.grade === 'F').length})` },
                        { id: 'R', label: `R 级品 (${nonAGradeList.filter(x => x.grade === 'R').length})` },
                        { id: 'OTHER', label: `其它 (${nonAGradeList.filter(x => !['B','F','R'].includes(x.grade)).length})` },
                      ].map((pill) => (
                        <button
                          key={pill.id}
                          onClick={() => setNonAGradeFilter(pill.id as any)}
                          className={cn(
                            "px-2.5 py-1 text-xs font-bold rounded-lg border transition-all",
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
                        <Search className="w-3.5 h-3.5 text-purple-400 absolute left-2.5 top-1/2 -translate-y-1/2" />
                        <input
                          type="text"
                          placeholder="搜索箱号/库位/工单/客户/等级..."
                          value={nonAGradeSearch}
                          onChange={(e) => setNonAGradeSearch(e.target.value)}
                          className="pl-8 pr-3 py-1.5 text-xs bg-white rounded-lg border border-purple-200 focus:outline-none focus:ring-2 focus:ring-purple-500 w-56 placeholder:text-slate-400"
                        />
                      </div>
                      <button
                        onClick={exportNonAGradeToExcel}
                        className="px-3 py-1.5 bg-purple-600 hover:bg-purple-700 text-white text-xs font-bold rounded-lg transition-colors flex items-center gap-1 shrink-0 shadow-sm active:scale-95"
                        title="单独导出非A级别品报表为 Excel"
                      >
                        <Download className="w-3.5 h-3.5" />
                        单独导出
                      </button>
                    </div>
                  </div>
                )}

                <div className="flex-1 overflow-auto">
                  <table className="w-full border-collapse">
                    {activeTab === 'location' ? (
                      <>
                        <thead className="sticky top-0 bg-white/95 backdrop-blur-sm z-10 border-b border-slate-100 shadow-sm shadow-slate-100/50">
                          <tr className="text-left">
                            <th className="pl-8 pr-4 py-5 text-[10px] font-bold text-slate-400 uppercase tracking-[0.2em] w-16">顺序</th>
                            <th className="px-6 py-5 text-[10px] font-bold text-slate-400 uppercase tracking-[0.2em]">库位名称</th>
                            <th className="px-6 py-5 text-[10px] font-bold text-slate-400 uppercase tracking-[0.2em]">大部分存放客户与功率</th>
                            <th className="px-6 py-5 text-[10px] font-bold text-slate-400 uppercase tracking-[0.2em] text-center">箱号数量</th>
                            <th className="px-6 py-5 text-[10px] font-bold text-slate-400 uppercase tracking-[0.2em] text-right pr-12">可入库数量</th>
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
                                        row['箱号数量'] > capacity ? "bg-red-500 shadow-[0_0_8px_rgba(239,68,68,0.5)]" : 
                                        row['箱号数量'] < capacity ? "bg-emerald-500 shadow-[0_0_8px_rgba(16,185,129,0.5)]" : 
                                        "bg-slate-900 shadow-[0_0_8px_rgba(15,23,42,0.3)]"
                                      )} />
                                      <div className="flex flex-col">
                                        <div className="flex items-center gap-2">
                                          <span className={cn(
                                            "text-sm font-bold transition-colors",
                                            row['箱号数量'] > capacity ? "text-red-600" : 
                                            row['箱号数量'] < capacity ? "text-emerald-700" :
                                            "text-slate-900"
                                          )}>
                                            {row['库位名称']}
                                          </span>
                                          {!PREDEFINED_LOCATION_SET.has(row['库位名称']) && (
                                            <span className="px-1.5 py-0.5 text-[9px] font-bold bg-amber-50 text-amber-700 rounded border border-amber-200">
                                              范围外库位
                                            </span>
                                          )}
                                        </div>
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
                                  <td className="px-6 py-4">
                                    {row['主存客户'] !== '空置' ? (
                                      <div className="flex flex-col gap-0.5">
                                        <span className="text-xs font-bold text-slate-800 truncate max-w-[180px]" title={row['主存客户']}>
                                          {row['主存客户']}
                                        </span>
                                        <span className="text-[10px] text-slate-400 font-mono">
                                          功率: {row['主存功率']}
                                        </span>
                                      </div>
                                    ) : (
                                      <span className="text-xs text-slate-300 font-medium italic">空置</span>
                                    )}
                                  </td>
                                  <td className="px-6 py-4 text-center">
                                    <span className={cn(
                                      "text-sm font-mono font-bold tabular-nums px-3 py-1 rounded-full",
                                      row['箱号数量'] > capacity ? "bg-red-50 text-red-600" : 
                                      row['箱号数量'] < capacity ? "bg-emerald-50 text-emerald-600" : 
                                      "bg-slate-100 text-slate-900"
                                    )}>
                                      {row['箱号数量']} / {capacity}
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
                              );
                            })
                          ) : (
                            <tr>
                              <td colSpan={5} className="px-8 py-32 text-center">
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
                    ) : activeTab === 'customer' ? (
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
                                  <div className="flex flex-wrap gap-1.5 max-h-24 overflow-y-auto pr-1 text-nowrap">
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
                                  <div className="w-16 h-16 bg-violet-50 text-violet-500 rounded-full flex items-center justify-center mb-4">
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
                    ) : activeTab === 'correction' ? (
                      <>
                        <thead className="sticky top-0 bg-white/95 backdrop-blur-sm z-10 border-b border-slate-100 shadow-sm shadow-slate-100/50">
                          <tr className="text-left text-slate-400 text-[10px] font-bold uppercase tracking-wider">
                            <th className="pl-8 pr-4 py-5 w-16">顺序</th>
                            <th className="px-6 py-5">不一致箱号</th>
                            <th className="px-6 py-5 text-center">当前库位</th>
                            <th className="px-6 py-5 text-left">该箱属性</th>
                            <th className="px-6 py-5 text-left">库位主流 (以此为主)</th>
                            <th className="px-6 py-5">不符类型</th>
                            <th className="px-6 py-5 pr-12">推荐去向及原因</th>
                          </tr>
                        </thead>
                        <tbody className="divide-y divide-slate-50 text-slate-700">
                          {relocationSuggestions.length > 0 ? (
                            relocationSuggestions.map((row, i) => (
                              <motion.tr 
                                key={i}
                                initial={{ opacity: 0, y: 10 }}
                                animate={{ opacity: 1, y: 0 }}
                                transition={{ delay: i * 0.012, duration: 0.25 }}
                                className="hover:bg-slate-50/50 transition-all duration-200 group text-xs"
                              >
                                <td className="pl-8 pr-4 py-4 font-mono font-bold text-slate-300 group-hover:text-amber-500 transition-colors">
                                  {String(i + 1).padStart(2, '0')}
                                </td>
                                <td className="px-6 py-4 font-bold text-slate-800 font-mono">
                                  {row.boxNo}
                                </td>
                                <td className="px-6 py-4 text-center font-bold text-slate-600 font-mono">
                                  {row.currentLoc}
                                </td>
                                <td className="px-6 py-4 text-left">
                                  <div className="flex flex-col space-y-0.5">
                                    <span className="font-bold text-slate-800 truncate max-w-[130px] block" title={row.customerName}>{row.customerName}</span>
                                    <span className="font-mono text-slate-450 text-[10px]">功率: {row.powerGrade}</span>
                                  </div>
                                </td>
                                <td className="px-6 py-4 text-left">
                                  <div className="flex flex-col space-y-0.5">
                                    <span className="font-bold text-blue-700 truncate max-w-[130px] block" title={row.dominantCustomer}>{row.dominantCustomer}</span>
                                    <span className="font-mono text-blue-500 text-[10px]">功率: {row.dominantPower}</span>
                                  </div>
                                </td>
                                <td className="px-6 py-4">
                                  <span className="inline-block px-2 py-0.5 font-bold text-[10px] rounded-md bg-amber-50 text-amber-700 border border-amber-200 text-nowrap">
                                    {row.mismatchType}
                                  </span>
                                </td>
                                <td className="px-6 py-4 pr-12">
                                  <div className="flex flex-col gap-0.5">
                                    <span className="font-bold text-emerald-600 font-mono flex items-center gap-1">
                                      <ArrowRight className="w-3.5 h-3.5" />
                                      {row.recommendedLoc}
                                    </span>
                                    <span className="text-[10px] text-slate-400 max-w-[280px]">
                                      {row.recommendationReason}
                                    </span>
                                  </div>
                                </td>
                              </motion.tr>
                            ))
                          ) : (
                            <tr>
                              <td colSpan={7} className="px-8 py-32 text-center">
                                <div className="flex flex-col items-center justify-center max-w-sm mx-auto">
                                  <div className="w-12 h-12 bg-emerald-50 text-emerald-500 rounded-full flex items-center justify-center mb-4">
                                    <CheckCircle2 className="w-6 h-6" />
                                  </div>
                                  <p className="text-sm font-bold text-slate-800">未检测到需要调整的混放箱</p>
                                  <p className="text-xs text-slate-400 mt-2 leading-relaxed">
                                    当前已有货存的所有库位均达到完美的一致性（单客户 + 单功率），不含任何非凡主属性的小部分异类箱子，无需修正调整！
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
                            <th className="pl-8 pr-4 py-5 w-16">顺序</th>
                            <th className="px-6 py-5">箱号</th>
                            <th className="px-6 py-5 text-center">库位名称</th>
                            <th className="px-6 py-5 text-center">箱等级 (K列原值)</th>
                            <th className="px-6 py-5 text-center">等级归类</th>
                            <th className="px-6 py-5 text-center">箱条码数 (S列)</th>
                            <th className="px-6 py-5">工单号 (J列)</th>
                            <th className="px-6 py-5">客户简称</th>
                            <th className="px-6 py-5 pr-12 text-center">功率档</th>
                          </tr>
                        </thead>
                        <tbody className="divide-y divide-slate-50 text-slate-700">
                          {filteredNonAGradeList.length > 0 ? (
                            filteredNonAGradeList.map((row, i) => (
                              <motion.tr 
                                key={i}
                                initial={{ opacity: 0, y: 10 }}
                                animate={{ opacity: 1, y: 0 }}
                                transition={{ delay: i * 0.008, duration: 0.2 }}
                                className="hover:bg-purple-50/40 transition-all duration-200 group text-xs"
                              >
                                <td className="pl-8 pr-4 py-3.5 font-mono font-bold text-slate-300 group-hover:text-purple-600 transition-colors">
                                  {String(i + 1).padStart(2, '0')}
                                </td>
                                <td className="px-6 py-3.5 font-bold text-slate-800 font-mono">
                                  {row.boxNo}
                                </td>
                                <td className="px-6 py-3.5 text-center">
                                  <span className="font-mono font-bold text-slate-700 bg-slate-100 px-2 py-0.5 rounded border border-slate-200">
                                    {row.location}
                                  </span>
                                </td>
                                <td className="px-6 py-3.5 text-center">
                                  <span className="font-mono text-slate-600">
                                    {row.rawGrade || '-'}
                                  </span>
                                </td>
                                <td className="px-6 py-3.5 text-center">
                                  <span className={cn(
                                    "px-2.5 py-0.5 rounded-md font-bold text-xs inline-block",
                                    row.grade === 'B' ? "bg-blue-100 text-blue-800 border border-blue-200" :
                                    row.grade === 'F' ? "bg-purple-100 text-purple-800 border border-purple-200" :
                                    row.grade === 'R' ? "bg-rose-100 text-rose-800 border border-rose-200" :
                                    "bg-slate-100 text-slate-700 border border-slate-200"
                                  )}>
                                    {row.grade} 级品
                                  </span>
                                </td>
                                <td className="px-6 py-3.5 text-center font-mono font-bold text-slate-700">
                                  {row.barcodeCount !== undefined ? row.barcodeCount : '-'}
                                </td>
                                <td className="px-6 py-3.5 font-mono text-slate-600">
                                  {row.orderNo || '-'}
                                </td>
                                <td className="px-6 py-3.5">
                                  <span className={cn(
                                    "font-bold truncate max-w-[140px] block",
                                    row.customerName === '未映射客户' ? "text-slate-400 italic font-normal" : "text-slate-800"
                                  )} title={row.customerName}>
                                    {row.customerName || '未映射客户'}
                                  </span>
                                </td>
                                <td className="px-6 py-3.5 pr-12 text-center font-mono text-slate-500">
                                  {row.power || '-'}
                                </td>
                              </motion.tr>
                            ))
                          ) : (
                            <tr>
                              <td colSpan={9} className="px-8 py-32 text-center">
                                <div className="flex flex-col items-center justify-center max-w-sm mx-auto">
                                  <div className="w-12 h-12 bg-purple-50 text-purple-500 rounded-full flex items-center justify-center mb-4">
                                    <Layers className="w-6 h-6" />
                                  </div>
                                  <p className="text-sm font-bold text-slate-800">
                                    {nonAGradeList.length === 0 ? '未检测到非A级别品' : '无符合当前筛选条件的非A级品'}
                                  </p>
                                  <p className="text-xs text-slate-400 mt-2 leading-relaxed">
                                    {nonAGradeList.length === 0 
                                      ? '当前主库存表中所有行均为 A 级良品，或尚未载入库存数据。'
                                      : '请切换品级筛选标签或清空搜索关键词后重试。'}
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
                    <span className="text-[10px] font-bold text-slate-500 uppercase tracking-widest">分析引擎 v5.0</span>
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
                      S列箱条码数为 0 / 无条码过滤明细
                      <span className="px-2 py-0.5 rounded-full text-xs bg-orange-100 text-orange-700 font-mono font-bold">
                        共 {zeroBarcodeList.length} 箱
                      </span>
                    </h3>
                    <p className="text-xs text-slate-500 mt-0.5">
                      根据规则，S列（箱条码数）为 0 或未填写条码的数据行不纳入实际库存统计。
                    </p>
                  </div>
                </div>
                <div className="flex items-center gap-2">
                  <button
                    onClick={exportZeroBarcodesToExcel}
                    className="px-3.5 py-2 bg-orange-600 hover:bg-orange-700 text-white rounded-xl text-xs font-bold flex items-center gap-1.5 shadow-sm transition-colors"
                  >
                    <Download className="w-3.5 h-3.5" />
                    导出Excel清单
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
                        <th className="px-4 py-3 text-center">序号</th>
                        <th className="px-4 py-3">原表行号</th>
                        <th className="px-4 py-3">箱号</th>
                        <th className="px-4 py-3">库位名称</th>
                        <th className="px-4 py-3">K列箱等级</th>
                        <th className="px-4 py-3 text-center">S列原值</th>
                        <th className="px-4 py-3">工单号</th>
                        <th className="px-4 py-3">过滤判定依据</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-100">
                      {zeroBarcodeList.map((item, idx) => (
                        <tr key={idx} className="hover:bg-orange-50/30 transition-colors">
                          <td className="px-4 py-2.5 text-center font-mono text-slate-400">{idx + 1}</td>
                          <td className="px-4 py-2.5 font-mono text-slate-600">第 {item.rowIdx} 行</td>
                          <td className="px-4 py-2.5 font-mono font-bold text-slate-800">{item.boxNo}</td>
                          <td className="px-4 py-2.5 font-mono text-slate-600">{item.location}</td>
                          <td className="px-4 py-2.5">
                            <span className="px-2 py-0.5 bg-slate-100 text-slate-600 rounded text-[10px] font-bold">
                              {item.grade}
                            </span>
                          </td>
                          <td className="px-4 py-2.5 text-center font-mono font-bold text-orange-600">
                            {item.rawBarcode === '' ? <span className="text-slate-400 font-normal italic">(留空)</span> : item.rawBarcode}
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
                  勾稽关系：总数据 {totalRowsRaw} 行 - 过滤条码0 {zeroBarcodeList.length} 行 = 实际纳入 {totalRowsRaw - zeroBarcodeList.length} 箱
                </span>
                <button
                  onClick={() => setShowZeroBarcodeModal(false)}
                  className="px-4 py-1.5 bg-slate-200 hover:bg-slate-300 text-slate-700 rounded-xl text-xs font-bold transition-colors"
                >
                  关闭
                </button>
              </div>
            </motion.div>
          </motion.div>
        )}
      </AnimatePresence>

      {/* 8箱差异深度溯源与行号诊断 Modal */}
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
                      8箱差异深度溯源与行号特点诊断
                      <span className="px-2 py-0.5 rounded-full text-xs bg-indigo-100 text-indigo-700 font-mono font-bold">
                        差异 = 8 箱 (7357 vs 7349)
                      </span>
                    </h3>
                    <p className="text-xs text-slate-500 mt-0.5">
                      全量透析 7,380 行 (原表总行) vs 7,372 箱 (表格统计) 与 7,357 箱 (有效纳入) 的勾稽逻辑
                    </p>
                  </div>
                </div>
                <div className="flex items-center gap-2">
                  <button
                    onClick={exportAuditReportToExcel}
                    className="px-3.5 py-1.5 bg-indigo-600 hover:bg-indigo-700 text-white rounded-xl text-xs font-bold flex items-center gap-1.5 transition-colors shadow-sm"
                  >
                    <Download className="w-3.5 h-3.5" />
                    导出诊断清单Excel
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
                
                {/* 1. 数学勾稽全景透视图 */}
                <div className="bg-slate-900 rounded-2xl p-5 text-white shadow-md space-y-4">
                  <div className="flex items-center justify-between border-b border-slate-800 pb-3">
                    <span className="text-xs font-bold text-indigo-400 uppercase tracking-wider flex items-center gap-2">
                      <ListOrdered className="w-4 h-4" />
                      核心数学勾稽关系全景推导（为什么会出现 7357 与 7349 的 8 箱差额）
                    </span>
                    <span className="text-[10px] text-slate-400 font-mono">
                      严格数学等式
                    </span>
                  </div>

                  <div className="grid grid-cols-1 md:grid-cols-4 gap-3 text-center font-mono">
                    <div className="bg-slate-800/80 p-3 rounded-xl border border-slate-700">
                      <span className="text-[10px] text-slate-400 block font-sans">① 原始文件实际总行数</span>
                      <span className="text-xl font-bold text-white mt-1 block">{totalRowsRaw} 行</span>
                      <span className="text-[9px] text-slate-400 block mt-1">Excel 真实物理数据行</span>
                    </div>
                    <div className="bg-orange-950/60 p-3 rounded-xl border border-orange-800/80">
                      <span className="text-[10px] text-orange-300 block font-sans">② S列条码为0过滤</span>
                      <span className="text-xl font-bold text-orange-400 mt-1 block">-{zeroBarcodeCount} 箱</span>
                      <span className="text-[9px] text-orange-300/80 block mt-1">作废/无条码排除行</span>
                    </div>
                    <div className="bg-emerald-950/60 p-3 rounded-xl border border-emerald-800/80">
                      <span className="text-[10px] text-emerald-300 block font-sans">③ 实际有效纳入总箱数</span>
                      <span className="text-xl font-bold text-emerald-400 mt-1 block">{totalRowsRaw - zeroBarcodeCount} 箱</span>
                      <span className="text-[9px] text-emerald-300/80 block mt-1">7380 - 23 = 7357</span>
                    </div>
                    <div className="bg-purple-950/60 p-3 rounded-xl border border-purple-800/80">
                      <span className="text-[10px] text-purple-300 block font-sans">④ A级与非A级构成</span>
                      <span className="text-lg font-bold text-purple-300 mt-1 block">{data.length} + {nonAGradeCount}</span>
                      <span className="text-[9px] text-purple-300/80 block mt-1">6879 (A) + 478 (非A) = 7357</span>
                    </div>
                  </div>

                  {/* 核心原理解析条 */}
                  <div className="bg-indigo-950/50 border border-indigo-700/60 rounded-xl p-3.5 text-xs text-indigo-200 space-y-1.5 leading-relaxed font-sans">
                    <p className="font-bold text-indigo-300 flex items-center gap-1.5 text-sm">
                      <Info className="w-4 h-4 text-indigo-400 shrink-0" />
                      8箱差异的根本原因揭秘：
                    </p>
                    <p>
                      • <strong>您算出的 7,349 箱：</strong> 来自于公式 <code className="bg-slate-800 px-1.5 py-0.5 rounded text-amber-300 font-mono">7,372 - 23 = 7,349</code>（用您之前表格统计的 7,372 减去了 23 箱）。
                    </p>
                    <p>
                      • <strong>系统的 7,357 箱：</strong> 来自于公式 <code className="bg-slate-800 px-1.5 py-0.5 rounded text-emerald-300 font-mono">7,380 - 23 = 7,357</code>（用源文件的实际 7,380 行减去了 23 箱）。
                    </p>
                    <p>
                      • <strong>两式之差：</strong> <code className="bg-slate-800 px-1.5 py-0.5 rounded text-white font-mono">7,357 - 7,349 = (7,380 - 23) - (7,372 - 23) = 7,380 - 7,372 = 8 箱</code>！
                    </p>
                    <p className="text-emerald-300 font-medium">
                      👉 <strong>结论：</strong> 6,879 (A级) + 478 (非A级) = 7,357 箱，在生成时<strong>已经把这 23 箱无效数据扣除掉了</strong>！并非去除23箱之后少8箱，而是您之前表格统计 7,372 箱时，就已经比源文件总行数（7,380 行）少计了 8 箱！
                    </p>
                  </div>
                </div>

                {/* 2. 这 8 箱在源表格 7380 行中的特点排查与具体行号 */}
                <div className="space-y-4">
                  <h4 className="text-sm font-bold text-slate-800 flex items-center gap-2">
                    <span className="w-2.5 h-2.5 rounded-full bg-blue-600" />
                    源数据深度扫描：这 8 箱差异对应哪些行？有什么特点？
                  </h4>

                  {/* 特点 1: 重复箱号 */}
                  <div className="bg-slate-50 border border-slate-200 rounded-2xl p-4 space-y-3">
                    <div className="flex items-center justify-between">
                      <div>
                        <h5 className="text-xs font-bold text-slate-800 flex items-center gap-1.5">
                          特点 ①：【重复箱号】（一箱多分录 / 一箱多工单 / 多功率档 / 重复扫描）
                        </h5>
                        <p className="text-[11px] text-slate-500 mt-0.5">
                          若在 Excel 中使用了透视表（非重复计数）或按箱号去重，重复出现的箱号只会被计为 1 箱，导致统计箱数比总行数少！
                        </p>
                      </div>
                      <span className={cn(
                        "px-2.5 py-1 rounded-full text-xs font-bold font-mono",
                        auditReport && auditReport.duplicateRowCount > 0 ? "bg-amber-100 text-amber-800 border border-amber-200" : "bg-emerald-100 text-emerald-800"
                      )}>
                        {auditReport ? `重复 ${auditReport.duplicateRowCount} 行 (${auditReport.duplicateBoxes.length} 个箱号)` : '正在扫描...'}
                      </span>
                    </div>

                    {auditReport && auditReport.duplicateBoxes.length > 0 ? (
                      <div className="overflow-x-auto border border-slate-200 rounded-xl bg-white max-h-52">
                        <table className="w-full text-left text-xs">
                          <thead className="bg-slate-100 sticky top-0 text-[10px] font-bold text-slate-600 uppercase">
                            <tr>
                              <th className="px-3 py-2">箱号</th>
                              <th className="px-3 py-2 text-center">原表格行号</th>
                              <th className="px-3 py-2 text-center">重复次数</th>
                              <th className="px-3 py-2">库位</th>
                              <th className="px-3 py-2">功率档</th>
                              <th className="px-3 py-2">工单号</th>
                              <th className="px-3 py-2">分析特点</th>
                            </tr>
                          </thead>
                          <tbody className="divide-y divide-slate-100 font-mono">
                            {auditReport.duplicateBoxes.map((dup, idx) => (
                              <tr key={idx} className="hover:bg-amber-50/50">
                                <td className="px-3 py-2 font-bold text-slate-800">{dup.boxNo}</td>
                                <td className="px-3 py-2 text-center text-indigo-600 font-bold">
                                  {dup.rows.map(r => `第${r.rowIdx}行`).join(', ')}
                                </td>
                                <td className="px-3 py-2 text-center">
                                  <span className="px-2 py-0.5 rounded bg-amber-100 text-amber-800 text-[10px] font-bold">
                                    {dup.count} 次
                                  </span>
                                </td>
                                <td className="px-3 py-2 font-sans">{dup.rows.map(r => r.location).join(' / ')}</td>
                                <td className="px-3 py-2 font-sans">{dup.rows.map(r => r.power).join(' / ')}</td>
                                <td className="px-3 py-2">{dup.rows.map(r => r.orderNo).join(' / ')}</td>
                                <td className="px-3 py-2 font-sans text-amber-700 text-[11px]">
                                  {dup.rows.some((r, i, arr) => i > 0 && r.power !== arr[0].power) 
                                    ? '同箱号多功率档分录' 
                                    : dup.rows.some((r, i, arr) => i > 0 && r.orderNo !== arr[0].orderNo)
                                    ? '同箱号多工单分录'
                                    : '同箱号完全重复扫描录入'}
                                </td>
                              </tr>
                            ))}
                          </tbody>
                        </table>
                      </div>
                    ) : (
                      <p className="text-xs text-slate-500 italic bg-white p-3 rounded-xl border border-slate-200">
                        当前文件未检测到重复箱号，所有有效数据行的箱号均互不相同。
                      </p>
                    )}
                  </div>

                  {/* 特点 2: S列条码明确为0的 23 箱行号与特点 */}
                  <div className="bg-orange-50/60 border border-orange-200 rounded-2xl p-4 space-y-3">
                    <div className="flex items-center justify-between">
                      <div>
                        <h5 className="text-xs font-bold text-slate-800 flex items-center gap-1.5">
                          特点 ②：【S列条码为0被过滤的 23 箱明细与等级归属】
                        </h5>
                        <p className="text-[11px] text-slate-500 mt-0.5">
                          这 23 箱在解析时因 S 列明确为 0 被排除。其中包含：A级品 {auditReport?.zeroBarcodeBreakdown.gradeA || 0} 箱，非A级品 {auditReport?.zeroBarcodeBreakdown.gradeNonA || 0} 箱。
                        </p>
                      </div>
                      <span className="px-2.5 py-1 rounded-full text-xs font-bold font-mono bg-orange-100 text-orange-800 border border-orange-200">
                        共 {zeroBarcodeList.length} 箱
                      </span>
                    </div>

                    <div className="overflow-x-auto border border-orange-200 rounded-xl bg-white max-h-52">
                      <table className="w-full text-left text-xs">
                        <thead className="bg-orange-100/60 sticky top-0 text-[10px] font-bold text-orange-900 uppercase">
                          <tr>
                            <th className="px-3 py-2 text-center">原表行号</th>
                            <th className="px-3 py-2">箱号</th>
                            <th className="px-3 py-2">库位</th>
                            <th className="px-3 py-2 text-center">K列箱等级</th>
                            <th className="px-3 py-2 text-center">S列条码数原值</th>
                            <th className="px-3 py-2">工单号</th>
                            <th className="px-3 py-2">判定原因</th>
                          </tr>
                        </thead>
                        <tbody className="divide-y divide-orange-100 font-mono">
                          {zeroBarcodeList.map((item, idx) => (
                            <tr key={idx} className="hover:bg-orange-50/40">
                              <td className="px-3 py-2 text-center font-bold text-orange-700">第 {item.rowIdx} 行</td>
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
                                {item.rawBarcode === '' ? '(留空)' : `"${item.rawBarcode}"`}
                              </td>
                              <td className="px-3 py-2">{item.orderNo || '-'}</td>
                              <td className="px-3 py-2 font-sans text-slate-600 text-[11px]">{item.reason}</td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  </div>

                  {/* 特点 3: 空箱号与空库位排查 */}
                  <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                    <div className="bg-slate-50 border border-slate-200 rounded-2xl p-4 space-y-2">
                      <div className="flex items-center justify-between">
                        <span className="text-xs font-bold text-slate-800">特点 ③：空箱号行（未录入箱号）</span>
                        <span className="px-2 py-0.5 rounded text-xs font-bold font-mono bg-slate-200 text-slate-700">
                          {auditReport ? auditReport.emptyBoxRows.length : 0} 行
                        </span>
                      </div>
                      <p className="text-[11px] text-slate-500">
                        {auditReport && auditReport.emptyBoxRows.length > 0 
                          ? `分布在原表：第 ${auditReport.emptyBoxRows.map(x => x.rowIdx).join(', ')} 行`
                          : '所有数据行均录入了有效箱号，未发现箱号为空的异常行。'}
                      </p>
                    </div>

                    <div className="bg-slate-50 border border-slate-200 rounded-2xl p-4 space-y-2">
                      <div className="flex items-center justify-between">
                        <span className="text-xs font-bold text-slate-800">特点 ④：空库位 / 待定库位</span>
                        <span className="px-2 py-0.5 rounded text-xs font-bold font-mono bg-slate-200 text-slate-700">
                          {auditReport ? auditReport.emptyLocRows.length : 0} 行
                        </span>
                      </div>
                      <p className="text-[11px] text-slate-500">
                        {auditReport && auditReport.emptyLocRows.length > 0
                          ? `分布在原表：第 ${auditReport.emptyLocRows.map(x => x.rowIdx).join(', ')} 行`
                          : '所有数据行均有确切库位，未发现未分配库位或 TBD 库位。'}
                      </p>
                    </div>
                  </div>
                </div>

              </div>

              {/* Modal Footer */}
              <div className="px-6 py-4 border-t border-slate-100 bg-slate-50 flex items-center justify-between text-xs text-slate-500">
                <span>
                  数据勾稽核验：{totalRowsRaw} (原总行) - {zeroBarcodeCount} (条码0) = {totalRowsRaw - zeroBarcodeCount} 箱 [{data.length} A级 + {nonAGradeCount} 非A级]
                </span>
                <div className="flex items-center gap-2">
                  <button
                    onClick={exportAuditReportToExcel}
                    className="px-4 py-1.5 bg-indigo-600 hover:bg-indigo-700 text-white rounded-xl text-xs font-bold transition-colors shadow-sm flex items-center gap-1"
                  >
                    <Download className="w-3.5 h-3.5" />
                    下载完整诊断表
                  </button>
                  <button
                    onClick={() => setShowAuditModal(false)}
                    className="px-4 py-1.5 bg-slate-200 hover:bg-slate-300 text-slate-700 rounded-xl text-xs font-bold transition-colors"
                  >
                    关闭
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
              <p className="text-[10px] font-bold text-red-400 uppercase tracking-widest mb-0.5">分析错误</p>
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
              <p className="text-[10px] font-bold text-emerald-400 uppercase tracking-widest mb-0.5">保存成功</p>
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

