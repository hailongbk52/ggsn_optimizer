import React, { useState, useEffect, useMemo, useRef } from "react";
import axios from "axios";
import * as XLSX from "xlsx";
import {
  FileSpreadsheet,
  Database,
  BarChart3,
  Settings,
  UploadCloud,
  Play,
  CheckCircle2,
  AlertCircle,
  TrendingUp,
  Download,
  AlertTriangle,
  RefreshCw,
  Table2,
  Filter,
  ChevronDown,
  ChevronUp,
  Edit2,
  Check,
  X,
  SlidersHorizontal,
  Maximize2,
  Minimize2,
} from "lucide-react";
import {
  BarChart,
  Bar,
  XAxis,
  YAxis,
  CartesianGrid,
  Tooltip,
  Legend,
  ResponsiveContainer,
  ReferenceLine,
} from "recharts";
import { evaluateGrid } from "./formulaEvaluator";
import type { GridRow, RefData, GlobalRefs, CellInfo } from "./formulaEvaluator";

const API_BASE = "http://127.0.0.1:8000/api";

// ── Node name blacklist ──────────────────────────────────────────────────────
const INVALID_NODE_PATTERNS = [
  /^GGSN_KV/i,
  /^GGSN_TQ/i,
  /^Back to/i,
  /gioi thieu/i,
];
function isValidNode(name: string) {
  return name && INVALID_NODE_PATTERNS.every((p) => !p.test(name.trim()));
}

type Tab = "sources" | "simulation" | "dashboard";

// ── Default mock headers (row1 group + row2 field name) ─────────────────────
const DEFAULT_HEADERS_ROW1: Record<string, string> = {
  E: "License node", F: "License node", G: "License VHKT", H: "License VHKT",
  J: "", K: "",
  L: "", M: "", N: "",
  O: "Tỷ lệ sử dụng", P: "Tỷ lệ sử dụng", Q: "Tỷ lệ sử dụng", O2: "Tỷ lệ sử dụng", O3: "Tỷ lệ sử dụng",
  T: "Tải GGSN sau khi cắt chuyển/cân tải", U: "Tải GGSN sau khi cắt chuyển/cân tải", V: "Tải GGSN sau khi cắt chuyển/cân tải",
  W: "Tải GGSN sau khi cắt chuyển/cân tải", X: "Tải GGSN sau khi cắt chuyển/cân tải",
};

const DEFAULT_HEADERS: Record<string, string> = {
  A: "STT", B: "Node", C: "Vendor", D: "Area",
  E: "License Bear", F: "License Throughput",
  G: "License Bear UCTT\n(110% License Node)", H: "License Throughput VHKT",
  I: "Total IPv4\n(v-internet)", I2: "Total IPv4\n(IMS)",
  J: "Bear sử dụng", K: "Throughput",
  L: "Bear Total sử dụng", M: "Bear IMS", N: "Throughput",
  O: "%Bear", P: "%Throughput", Q: "IMS",
  R: "Tải IMS Trước cắt", O2: "% Tải IMS trước", S: "Tải IMS Sau cắt", O3: "% Tải IMS sau",
  T: "Bear sử dụng", U: "Throughput dự kiến", V: "Bearer IMS dự kiến",
  W: "%Bear", X: "%Throughput",
  AH: "Weight", AJ: "ON/OFF", AK: "NEW WEIGHT", AL: "Change",
  AM: "HW/NFVI Site", AN: "Tăng/giảm",
};

// ── Table templates for individual input mode ────────────────────────────────
const TABLE_TEMPLATES = {
  license: {
    label: "Bảng License GGSN",
    description: "Thông tin license của từng GGSN node",
    columns: ["Node", "Vendor", "Area", "License Bear", "License Throughput", "License Bear UCTT (110%)", "License Throughput VHKT"],
    sqlHint: "SELECT node, vendor, area, lic_bear, lic_throughput, lic_bear_uctt, lic_throughput_vhkt FROM table_license;",
    rows: [{ Node: "GGPD04", Vendor: "Huawei", Area: "KV1", "License Bear": 2500000, "License Throughput": 100000, "License Bear UCTT (110%)": 2750000, "License Throughput VHKT": 110000 }],
  },
  current: {
    label: "Bảng Tải Hiện Tại",
    description: "Số liệu tải thực tế hiện tại (Bear sử dụng, Throughput…)",
    columns: ["Node", "Vendor", "Bear sử dụng", "Throughput", "Bear Total sử dụng", "Bear IMS", "Total IPv4 (v-internet)", "Total IPv4 (IMS)"],
    sqlHint: "SELECT node, vendor, bear_su_dung, throughput, bear_total_su_dung, bear_ims, ipv4_internet, ipv4_ims FROM table_current;",
    rows: [{ Node: "GGPD04", Vendor: "Huawei", "Bear sử dụng": 2333754, "Throughput": 88777.52, "Bear Total sử dụng": 2333754, "Bear IMS": 1166877, "Total IPv4 (v-internet)": 2000000, "Total IPv4 (IMS)": 300000 }],
  },
  weight: {
    label: "Bảng Trọng Số Cân Tải",
    description: "Weight hiện tại và Weight mới để tính toán cân tải",
    columns: ["Node", "Vendor", "Weight (AH)", "NEW WEIGHT (AK)", "ON/OFF"],
    sqlHint: "SELECT node, vendor, weight, new_weight, on_off FROM table_weight;",
    rows: [{ Node: "GGPD04", Vendor: "Huawei", "Weight (AH)": 90, "NEW WEIGHT (AK)": 90, "ON/OFF": 1 }],
  },
  ims_routing: {
    label: "Bảng IMS Routing",
    description: "Mapping IMS site cho từng node → cột IMS(Q) trong bảng mô phỏng. Mặc định: \"IMS\"",
    columns: ["Node", "Vendor", "IMS_site"],
    sqlHint: "SELECT node, vendor, ims_site FROM table_ims_routing;",
    rows: [{ Node: "GGPD04", Vendor: "Huawei", IMS_site: "IMS" }],
  },
  hw_site: {
    label: "Bảng HW/NFVI Site",
    description: "Mapping HW/NFVI Site cho từng node → cột HW/NFVI Site(AM). Mặc định: \"site\"",
    columns: ["Node", "Vendor", "HW/NFVI Site"],
    sqlHint: "SELECT node, vendor, hw_nfvi_site FROM table_hw_site;",
    rows: [{ Node: "GGPD04", Vendor: "Huawei", "HW/NFVI Site": "site" }],
  },
  ims_license: {
    label: "Bảng License IMS",
    description: "License IMS cho từng hệ thống IMS → dùng để tính % Tải IMS. Mặc định: 5000000",
    columns: ["Node", "Vendor", "License"],
    sqlHint: "SELECT node, vendor, license FROM table_ims_license;",
    rows: [{ Node: "IMS_HN", Vendor: "Huawei", License: 5000000 }],
  },
};

// ── Helpers ──────────────────────────────────────────────────────────────────
function pct(v: any) {
  if (typeof v === "number") return `${(v * 100).toFixed(1)}%`;
  return v ?? "";
}
function numFmt(v: any) {
  if (typeof v === "number") return v.toLocaleString(undefined, { maximumFractionDigits: 2 });
  return v ?? "";
}

const PCT_COLS = new Set(["O", "P", "W", "X", "AE", "AF", "AG"]);
const IMS_PCT_COLS = new Set(["O2", "O3"]);

export default function App() {
  const [activeTab, setActiveTab] = useState<Tab>("simulation");
  const [isFullView, setIsFullView] = useState(false);

  // Grid data
  const [fileName, setFileName] = useState("Phòng Di động - TT VHKTTC");
  const [headersRow1, setHeadersRow1] = useState<Record<string, string>>(DEFAULT_HEADERS_ROW1);
  const [headers, setHeaders] = useState<Record<string, string>>(DEFAULT_HEADERS);
  const [rows, setRows] = useState<GridRow[]>([]);
  const [refData, setRefData] = useState<RefData>({});
  const [globalRefs, setGlobalRefs] = useState<GlobalRefs>({});
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);

  // Column name editing
  const [editingHeader, setEditingHeader] = useState<string | null>(null);
  const [tempHeaderVal, setTempHeaderVal] = useState("");

  // Cell editing
  const [editingCell, setEditingCell] = useState<{ rowIdx: number; colLetter: string } | null>(null);
  const [tempFormula, setTempFormula] = useState("");
  const [tempValue, setTempValue] = useState("");

  // DB state
  const [connStr, setConnStr] = useState("sqlite://mock_internal_ggsn.db");
  const [dbQueryMap, setDbQueryMap] = useState<Record<string, string>>(
    Object.fromEntries(Object.keys(TABLE_TEMPLATES).map((k) => [k, TABLE_TEMPLATES[k as keyof typeof TABLE_TEMPLATES].sqlHint]))
  );
  const [dbResults, setDbResults] = useState<Record<string, any>>({});
  const [dbStatuses, setDbStatuses] = useState<Record<string, "idle" | "success" | "error">>({});

  // Submenu and Pagination State
  const [sourcesExpanded, setSourcesExpanded] = useState(true);
  const [currentSourceTab, setCurrentSourceTab] = useState<"license" | "current" | "weight" | "ims_routing" | "hw_site" | "ims_license">("license");
  const [tablePages, setTablePages] = useState<Record<string, number>>({
    license: 1, current: 1, weight: 1, ims_routing: 1, hw_site: 1, ims_license: 1
  });
  const [tablePageSize, setTablePageSize] = useState(25);

  // Multi-cell selection state
  const [selectedCells, setSelectedCells] = useState<Set<string>>(new Set());
  const [anchorCell, setAnchorCell] = useState<{ rowIdx: number; colLetter: string } | null>(null);

  // Individual table input data
  const [tableInputData, setTableInputData] = useState<Record<string, any[]>>(
    Object.fromEntries(Object.keys(TABLE_TEMPLATES).map((k) => [k, TABLE_TEMPLATES[k as keyof typeof TABLE_TEMPLATES].rows.map(r => ({ ...r }))]))
  );

  // Threshold & sorting
  const [overloadThreshold, setOverloadThreshold] = useState(90);
  const [sortAsc, setSortAsc] = useState(true);

  // Table zoom and filter states
  const [zoomScale, setZoomScale] = useState(100);
  const [columnFilters, setColumnFilters] = useState<Record<string, string>>({});

  // SQL Scheduler states
  const [schedules, setSchedules] = useState<any[]>([]);
  const [scheduleTypes, setScheduleTypes] = useState<Record<string, string>>({});
  const [scheduleActives, setScheduleActives] = useState<Record<string, boolean>>({});

  // Dashboard filter
  const [selectedNodes, setSelectedNodes] = useState<string[]>([]);
  const [selectedAreas, setSelectedAreas] = useState<string[]>([]);
  const [selectedHwSites, setSelectedHwSites] = useState<string[]>([]);
  const [nodeFilterOpen, setNodeFilterOpen] = useState(false);
  const [nodeSearch, setNodeSearch] = useState("");

  // Single SQL db state
  const [singleConnStr, setSingleConnStr] = useState("sqlite://mock_internal_ggsn.db");
  const [singleSql, setSingleSql] = useState("SELECT * FROM ggsn_metrics;");
  const [singleDbResult, setSingleDbResult] = useState<any>(null);
  const [singleDbStatus, setSingleDbStatus] = useState<"idle" | "success" | "error">("idle");
  const [mappingCol, setMappingCol] = useState("J");

  const saveTableToBackend = async (key: string, dataToSave: any[]) => {
    try {
      await axios.post(`${API_BASE}/save-table-data`, {
        table_key: key,
        rows: dataToSave
      });
    } catch (err) {
      console.error("Lưu dữ liệu bảng thất bại:", err);
    }
  };

  const fetchSchedules = async () => {
    try {
      const res = await axios.get(`${API_BASE}/get-schedules`);
      if (res.data.success) {
        setSchedules(res.data.schedules);
        const types: Record<string, string> = {};
        const actives: Record<string, boolean> = {};
        res.data.schedules.forEach((s: any) => {
          types[s.table_key] = s.schedule_type;
          actives[s.table_key] = s.is_active === 1;
        });
        setScheduleTypes(types);
        setScheduleActives(actives);
      }
    } catch (err) {
      console.error("Failed to load SQL schedules", err);
    }
  };

  const saveSchedule = async (tableKey: string) => {
    try {
      const query = dbQueryMap[tableKey] || TABLE_TEMPLATES[tableKey as keyof typeof TABLE_TEMPLATES].sqlHint;
      const type = scheduleTypes[tableKey] || "manual";
      const active = scheduleActives[tableKey] ? 1 : 0;
      
      const res = await axios.post(`${API_BASE}/save-schedule`, {
        table_key: tableKey,
        query: query,
        schedule_type: type,
        is_active: active
      });
      if (res.data.success) {
        alert("Lưu cấu hình lập lịch SQL thành công!");
        fetchSchedules();
      }
    } catch (err) {
      alert("Lưu lập lịch thất bại: " + String(err));
    }
  };

  const fetchDbData = async () => {
    setLoading(true);
    try {
      const res = await axios.get(`${API_BASE}/get-all-data`);
      if (res.data.success) {
        const loadedTables = {
          license: res.data.license || [],
          current: res.data.current || [],
          weight: res.data.weight || [],
          ims_routing: res.data.ims_routing || [],
          hw_site: res.data.hw_site || [],
          ims_license: res.data.ims_license || [],
        };
        setTableInputData(loadedTables);
        rebuildSimulationGrid(loadedTables);
      } else {
        generateClientMockData();
      }
    } catch (err) {
      console.error("Failed to load from database. Fallback to mock.", err);
      generateClientMockData();
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    fetchDbData();
    fetchSchedules();
  }, []);

  // Rebuild the main simulation grid (rows) from the 3 tables
  function rebuildSimulationGrid(currentTables = tableInputData) {
    const allNodesSet = new Set<string>();
    const addNodes = (list: any[]) => {
      if (Array.isArray(list)) {
        list.forEach(r => {
          const nodeName = String(r.Node || r.node || r["GGSN Node"] || "").trim().toUpperCase();
          if (nodeName && isValidNode(nodeName)) {
            allNodesSet.add(nodeName);
          }
        });
      }
    };
    addNodes(currentTables.license);
    addNodes(currentTables.current);
    addNodes(currentTables.weight);
    const uniqueNodes = Array.from(allNodesSet).sort();

    if (uniqueNodes.length === 0) {
      setRows([]);
      return;
    }

    const nextRows: GridRow[] = uniqueNodes.map((nodeName, index) => {
      const licMatch = currentTables.license?.find((r: any) => String(r.Node || r.node || "").trim().toUpperCase() === nodeName);
      const curMatch = currentTables.current?.find((r: any) => String(r.Node || r.node || "").trim().toUpperCase() === nodeName);
      const wtMatch = currentTables.weight?.find((r: any) => String(r.Node || r.node || "").trim().toUpperCase() === nodeName);
      const imsMatch = currentTables.ims_routing?.find((r: any) => String(r.Node || r.node || "").trim().toUpperCase() === nodeName);
      const hwMatch = currentTables.hw_site?.find((r: any) => String(r.Node || r.node || "").trim().toUpperCase() === nodeName);

      const vendor = String(licMatch?.Vendor || licMatch?.vendor || curMatch?.Vendor || curMatch?.vendor || wtMatch?.Vendor || wtMatch?.vendor || "Huawei");
      const imsVal = String(imsMatch?.IMS_site || imsMatch?.ims_site || "IMS");
      const hwVal = String(hwMatch?.["HW/NFVI Site"] || hwMatch?.hw_nfvi_site || "site");

      // Find IMS license for this node's IMS system
      const imsLicMatch = currentTables.ims_license?.find((r: any) => String(r.Node || r.node || "").trim().toUpperCase() === imsVal.trim().toUpperCase());
      const imsLicVal = Number(imsLicMatch?.License || imsLicMatch?.license || 0);

      const g = Number(licMatch?.["License Bear UCTT (110%)"] || licMatch?.lic_bear_uctt || licMatch?.["License Bear UCTT"] || 0);
      const h = Number(licMatch?.["License Throughput VHKT"] || licMatch?.lic_throughput_vhkt || 0);
      const l = Number(curMatch?.["Bear Total sử dụng"] || curMatch?.bear_total_su_dung || curMatch?.bear_total || 0);
      const m = Number(curMatch?.["Bear IMS"] || curMatch?.bear_ims || (l * 0.5));
      const n = Number(curMatch?.["Throughput"] || curMatch?.throughput || 0);

      const ah = Number(wtMatch?.["Weight (AH)"] || wtMatch?.weight || wtMatch?.Weight || 0);
      const ak = wtMatch ? Number(wtMatch?.["NEW WEIGHT (AK)"] || wtMatch?.new_weight || wtMatch?.["NEW WEIGHT"] || 0) : ah;
      const aj = wtMatch ? Number((wtMatch?.["ON/OFF"] || wtMatch?.on_off || wtMatch?.is_active) ?? 1) : 1;

      return {
        row_num: index + 3,
        A: { header: "STT", value: index + 1, formula: "", is_formula: false },
        B: { header: "Node", value: nodeName, formula: "", is_formula: false },
        C: { header: "Vendor", value: vendor, formula: "", is_formula: false },
        D: { header: "Area", value: String(licMatch?.Area || licMatch?.area || ""), formula: "", is_formula: false },
        E: { header: "License Bear", value: Number(licMatch?.["License Bear"] || licMatch?.lic_bear || 0), formula: "", is_formula: false },
        F: { header: "License Throughput", value: Number(licMatch?.["License Throughput"] || licMatch?.lic_throughput || 0), formula: "", is_formula: false },
        G: { header: "License Bear UCTT", value: g, formula: "", is_formula: false },
        H: { header: "License Throughput VHKT", value: h, formula: "", is_formula: false },
        I: { header: "Total IPv4 (v-internet)", value: Number(curMatch?.["Total IPv4 (v-internet)"] || curMatch?.ipv4_internet || 0), formula: "", is_formula: false },
        I2: { header: "Total IPv4 (IMS)", value: Number(curMatch?.["Total IPv4 (IMS)"] || curMatch?.ipv4_ims || 0), formula: "", is_formula: false },
        J: { header: "Bear sử dụng", value: Number(curMatch?.["Bear sử dụng"] || curMatch?.bear_su_dung || 0), formula: "", is_formula: false },
        K: { header: "Throughput", value: n, formula: "", is_formula: false },
        L: { header: "Bear Total sử dụng", value: l, formula: "", is_formula: false },
        M: { header: "Bear IMS", value: m, formula: "=L*50%", is_formula: true },
        N: { header: "Throughput", value: n, formula: "", is_formula: false },
        O: { header: "%Bear", value: g > 0 ? l / g : 0, formula: "=L/G", is_formula: true },
        P: { header: "%Throughput", value: h > 0 ? n / h : 0, formula: "=N/H", is_formula: true },
        Q: { header: "IMS", value: imsVal, formula: "", is_formula: false },
        R: { header: "Tải IMS Trước cắt", value: m, formula: "", is_formula: false },
        O2: { header: "% Tải IMS trước", value: imsLicVal > 0 ? m / imsLicVal : 0, formula: "=R/IMS_LIC", is_formula: true },
        S: { header: "Tải IMS Sau cắt", value: m, formula: "", is_formula: false },
        O3: { header: "% Tải IMS sau", value: imsLicVal > 0 ? m / imsLicVal : 0, formula: "=S/IMS_LIC", is_formula: true },
        T: { header: "Bear sử dụng DK", value: l, formula: "=(L+(AK/ΣAKAK-AH/ΣAHAH)*ΣL)*AJ", is_formula: true },
        U: { header: "Throughput dự kiến", value: n, formula: "=(N+(AK/ΣAKAK-AH/ΣAHAH)*ΣN)*AJ", is_formula: true },
        V: { header: "Bearer IMS dự kiến", value: m, formula: "=M+(AK/ΣAKAK-AH/ΣAHAH)*ΣM", is_formula: true },
        W: { header: "%Bear DK", value: g > 0 ? l / g : 0, formula: "=T/G", is_formula: true },
        X: { header: "%Throughput DK", value: h > 0 ? n / h : 0, formula: "=U/H", is_formula: true },
        AH: { header: "Weight", value: ah, formula: "", is_formula: false },
        AJ: { header: "ON/OFF", value: aj, formula: "=IF(AK=0,0,1)", is_formula: true },
        AK: { header: "NEW WEIGHT", value: ak, formula: "", is_formula: false },
        AL: { header: "Change", value: ah === ak ? "NO" : "YES", formula: "=IF(AH=AK,NO,YES)", is_formula: true },
        AM: { header: "HW/NFVI Site", value: hwVal, formula: "", is_formula: false },
        AN: { header: "Tăng/giảm", value: 0, formula: "=U-N", is_formula: true },
        _imsLic: imsLicVal,
      };
    });

    setRows(recalcRows(nextRows));
  }

  // ── Mock data ──────────────────────────────────────────────────────────────
  function generateClientMockData() {
    setHeaders(DEFAULT_HEADERS);
    setHeadersRow1(DEFAULT_HEADERS_ROW1);

    const nodes = [
      { id: 3, name: "GGPD04", vendor: "Huawei", g: 2750000, h: 110000, i: 2000000, i2: 350000, j: 2333754, k: 88777.52, l: 2333754, n: 88777, ah: 90, ak: 90 },
      { id: 4, name: "GGPD05", vendor: "Huawei", g: 2750000, h: 110000, i: 2000000, i2: 340000, j: 2306341, k: 88791.13, l: 2306341, n: 88791, ah: 90, ak: 90 },
      { id: 5, name: "GGPD06", vendor: "Huawei", g: 3850000, h: 130000, i: 3000000, i2: 480000, j: 3075825, k: 116824.07, l: 3075825, n: 116824, ah: 120, ak: 120 },
      { id: 6, name: "GGPD07", vendor: "Huawei", g: 3850000, h: 130000, i: 3000000, i2: 470000, j: 3048884, k: 116981.69, l: 3048884, n: 116981, ah: 120, ak: 120 },
      { id: 7, name: "GGHL04", vendor: "ZTE", g: 1200000, h: 50000, i: 1000000, i2: 200000, j: 1030469, k: 40625.70, l: 1030469, n: 40625, ah: 60, ak: 60 },
      { id: 8, name: "GGHL05", vendor: "ZTE", g: 1200000, h: 50000, i: 1000000, i2: 180000, j: 1037096, k: 40537.58, l: 1037096, n: 40537, ah: 60, ak: 60 },
      { id: 9, name: "GGHL06", vendor: "ZTE", g: 1200000, h: 50000, i: 1000000, i2: 190000, j: 299166, k: 3101.05, l: 299166, n: 3101, ah: 60, ak: 60 },
      { id: 10, name: "GGHL07", vendor: "ZTE", g: 1200000, h: 50000, i: 1000000, i2: 200000, j: 1118598, k: 42880.41, l: 1118598, n: 42880, ah: 60, ak: 60 },
      { id: 11, name: "GGHL11", vendor: "Ericsson", g: 2750000, h: 110000, i: 2000000, i2: 360000, j: 2490029, k: 89261.57, l: 2490029, n: 89261, ah: 90, ak: 90 },
      { id: 12, name: "GGHL12", vendor: "Ericsson", g: 2750000, h: 110000, i: 2000000, i2: 340000, j: 2474712, k: 87932.89, l: 2474712, n: 87932, ah: 90, ak: 90 },
    ];

    const license = nodes.map(nd => ({
      Node: nd.name,
      Vendor: nd.vendor,
      Area: nd.name.startsWith("GGPD") ? "KV1" : nd.name.startsWith("GGHL0") ? "KV2" : "KV3",
      "License Bear": nd.g / 1.1,
      "License Throughput": nd.h / 1.1,
      "License Bear UCTT (110%)": nd.g,
      "License Throughput VHKT": nd.h
    }));

    const current = nodes.map(nd => ({
      Node: nd.name,
      Vendor: nd.vendor,
      "Bear sử dụng": nd.j,
      Throughput: nd.k,
      "Bear Total sử dụng": nd.l,
      "Bear IMS": nd.l * 0.5,
      "Total IPv4 (v-internet)": nd.i,
      "Total IPv4 (IMS)": nd.i2
    }));

    const weight = nodes.map(nd => ({
      Node: nd.name,
      Vendor: nd.vendor,
      "Weight (AH)": nd.ah,
      "NEW WEIGHT (AK)": nd.ak,
      "ON/OFF": 1
    }));

    const mockedTables = { license, current, weight, ims_license: [
      { Node: "IMS_HN", Vendor: "Huawei", License: 5000000 },
      { Node: "IMS_HCM", Vendor: "Ericsson", License: 4000000 },
    ]};
    setTableInputData(mockedTables);
    rebuildSimulationGrid(mockedTables);

    const mockRows: GridRow[] = nodes.map(nd => {
      const m = nd.l * 0.5;
      const o = nd.g > 0 ? nd.l / nd.g : 0;
      const p = nd.h > 0 ? nd.n / nd.h : 0;
      const sumAK = nodes.reduce((s, x) => s + x.ak, 0);
      const sumAH = nodes.reduce((s, x) => s + x.ah, 0);
      const sumL = nodes.reduce((s, x) => s + x.l, 0);
      const sumN = nodes.reduce((s, x) => s + x.n, 0);
      const aj = nd.ak !== 0 ? 1 : 0;
      const t = (nd.l + (nd.ak / sumAK - nd.ah / sumAH) * sumL) * aj;
      const u = (nd.n + (nd.ak / sumAK - nd.ah / sumAH) * sumN) * aj;
      const v = m + (nd.ak / sumAK - nd.ah / sumAH) * nodes.reduce((s, x) => s + x.l * 0.5, 0);
      const w = nd.g > 0 ? t / nd.g : 0;
      const x = nd.h > 0 ? u / nd.h : 0;

      return {
        row_num: nd.id,
        A: { header: "STT", value: nd.id - 2, formula: "", is_formula: false },
        B: { header: "Node", value: nd.name, formula: "", is_formula: false },
        C: { header: "Vendor", value: "Huawei", formula: "", is_formula: false },
        D: { header: "Area", value: "KV1", formula: "", is_formula: false },
        E: { header: "License Bear", value: nd.g / 1.1, formula: "", is_formula: false },
        F: { header: "License Throughput", value: nd.h / 1.1, formula: "", is_formula: false },
        G: { header: "License Bear UCTT", value: nd.g, formula: "", is_formula: false },
        H: { header: "License Throughput VHKT", value: nd.h, formula: "", is_formula: false },
        I: { header: "Total IPv4 (v-internet)", value: nd.i, formula: "", is_formula: false },
        I2: { header: "Total IPv4 (IMS)", value: nd.i2, formula: "", is_formula: false },
        J: { header: "Bear sử dụng", value: nd.j, formula: "", is_formula: false },
        K: { header: "Throughput", value: nd.k, formula: "", is_formula: false },
        L: { header: "Bear Total sử dụng", value: nd.l, formula: "", is_formula: false },
        M: { header: "Bear IMS", value: m, formula: "=L*50%", is_formula: true },
        N: { header: "Throughput", value: nd.n, formula: "", is_formula: false },
        O: { header: "%Bear", value: o, formula: "=L/G", is_formula: true },
        P: { header: "%Throughput", value: p, formula: "=N/H", is_formula: true },
        Q: { header: "IMS", value: "IMS_HN", formula: "", is_formula: false },
        R: { header: "Tải IMS Trước cắt", value: m, formula: "", is_formula: false },
        O2: { header: "% Tải IMS trước", value: 5000000 > 0 ? m / 5000000 : 0, formula: "=R/IMS_LIC", is_formula: true },
        S: { header: "Tải IMS Sau cắt", value: v, formula: "", is_formula: false },
        O3: { header: "% Tải IMS sau", value: 5000000 > 0 ? v / 5000000 : 0, formula: "=S/IMS_LIC", is_formula: true },
        T: { header: "Bear sử dụng DK", value: t, formula: "=(L+(AK/ΣAKAK-AH/ΣAHAH)*ΣL)*AJ", is_formula: true },
        U: { header: "Throughput dự kiến", value: u, formula: "=(N+(AK/ΣAKAK-AH/ΣAHAH)*ΣN)*AJ", is_formula: true },
        V: { header: "Bearer IMS dự kiến", value: v, formula: "=M+(AK/ΣAKAK-AH/ΣAHAH)*ΣM", is_formula: true },
        W: { header: "%Bear DK", value: w, formula: "=T/G", is_formula: true },
        X: { header: "%Throughput DK", value: x, formula: "=U/H", is_formula: true },
        AH: { header: "Weight", value: nd.ah, formula: "", is_formula: false },
        AJ: { header: "ON/OFF", value: aj, formula: "=IF(AK=0,0,1)", is_formula: true },
        AK: { header: "NEW WEIGHT", value: nd.ak, formula: "", is_formula: false },
        AL: { header: "Change", value: nd.ah === nd.ak ? "NO" : "YES", formula: "=IF(AH=AK,NO,YES)", is_formula: true },
        AM: { header: "HW/NFVI Site", value: "HNI", formula: "", is_formula: false },
        AN: { header: "Tăng/giảm", value: u - nd.n, formula: "=U-N", is_formula: true },
        _imsLic: 5000000,
      };
    });

    setRows(mockRows);
  }

  // ── Filtered & sorted rows ─────────────────────────────────────────────────
  const displayRows = useMemo(() => {
    const validRows = rows.filter(r => {
      const node = (r["B"] as CellInfo)?.value;
      if (!(typeof node === "string" && isValidNode(node))) return false;
      
      return Object.keys(columnFilters).every(col => {
        const query = columnFilters[col]?.trim().toLowerCase();
        if (!query) return true;
        const cell = r[col] as CellInfo;
        if (!cell) return false;
        let valStr = "";
        const val = cell.value;
        if (PCT_COLS.has(col) && typeof val === "number") {
          valStr = `${(val * 100).toFixed(1)}%`;
        } else if (typeof val === "number") {
          valStr = val.toLocaleString(undefined, { maximumFractionDigits: 2 });
        } else {
          valStr = String(val ?? "");
        }
        return valStr.toLowerCase().includes(query);
      });
    });
    return [...validRows].sort((a, b) => {
      const na = String((a["B"] as CellInfo)?.value || "");
      const nb = String((b["B"] as CellInfo)?.value || "");
      return sortAsc ? na.localeCompare(nb) : nb.localeCompare(na);
    });
  }, [rows, sortAsc, columnFilters]);

  const allNodeNames = useMemo(() =>
    displayRows.map(r => String((r["B"] as CellInfo)?.value || "")), [displayRows]);

  // ── Column stats for highlighting ─────────────────────────────────────────
  const colStats = useMemo(() => {
    const stats: Record<string, { topSet: Set<number>; overThreshold: boolean }> = {};
    const colKeys = Object.keys(headers);

    colKeys.forEach(col => {
      const values = displayRows
        .map(r => (r[col] as CellInfo)?.value)
        .filter(v => typeof v === "number") as number[];

      if (values.length === 0) return;

      if (PCT_COLS.has(col) || IMS_PCT_COLS.has(col)) {
        const overThresh = values.filter(v => v * 100 > overloadThreshold);
        if (overThresh.length > 0) {
          stats[col] = { topSet: new Set(overThresh), overThreshold: true };
        } else {
          const sorted = [...values].sort((a, b) => b - a);
          const top5 = new Set(sorted.slice(0, 5));
          stats[col] = { topSet: top5, overThreshold: false };
        }
      } else {
        const sorted = [...values].sort((a, b) => b - a);
        const top5 = new Set(sorted.slice(0, 5));
        stats[col] = { topSet: top5, overThreshold: false };
      }
    });
    return stats;
  }, [displayRows, headers, overloadThreshold]);

  function getCellHighlight(col: string, value: any): string {
    if (typeof value !== "number") return "";
    const stat = colStats[col];
    if (!stat) return "";

    if (PCT_COLS.has(col) || IMS_PCT_COLS.has(col)) {
      if (stat.overThreshold && value * 100 > overloadThreshold) return "bg-red-500/20 text-red-300 font-bold";
      if (!stat.overThreshold && stat.topSet.has(value)) return "bg-amber-500/20 text-amber-300 font-bold";
    } else {
      if (stat.topSet.has(value)) return "bg-blue-500/15 text-blue-200 font-semibold";
    }
    return "";
  }

  // ── recalc after weight edit ───────────────────────────────────────────────
  function recalcRows(inputRows: GridRow[]) {
    const sumAK = inputRows.reduce((s, r) => {
      const v = (r["AK"] as CellInfo)?.value;
      return s + (typeof v === "number" ? v : 0);
    }, 0);
    const sumAH = inputRows.reduce((s, r) => {
      const v = (r["AH"] as CellInfo)?.value;
      return s + (typeof v === "number" ? v : 0);
    }, 0);
    const sumL = inputRows.reduce((s, r) => {
      const v = (r["L"] as CellInfo)?.value;
      return s + (typeof v === "number" ? v : 0);
    }, 0);
    const sumN = inputRows.reduce((s, r) => {
      const v = (r["N"] as CellInfo)?.value;
      return s + (typeof v === "number" ? v : 0);
    }, 0);
    const sumM = inputRows.reduce((s, r) => {
      const v = (r["M"] as CellInfo)?.value;
      return s + (typeof v === "number" ? v : 0);
    }, 0);

    // Pre-compute IMS group sums for columns R and S
    // R = sum of M for all rows sharing the same Q (IMS system name)
    // S = sum of V for all rows sharing the same Q (IMS system name)
    // We first do a pass to accumulate per-group sums, then assign
    const imsSumM: Record<string, number> = {};
    const imsSumV_pre: Record<string, number> = {};
    inputRows.forEach(r => {
      const q = String((r["Q"] as CellInfo)?.value || "");
      const m = (r["M"] as CellInfo)?.value as number || 0;
      imsSumM[q] = (imsSumM[q] || 0) + m;
      // V before recalc: compute delta here to get V_new
      const ak = (r["AK"] as CellInfo)?.value as number || 0;
      const ah = (r["AH"] as CellInfo)?.value as number || 0;
      const aj = ak !== 0 ? 1 : 0;
      const delta = sumAK > 0 ? (ak / sumAK - ah / sumAH) : 0;
      const v_new = (m + delta * sumM) * aj;
      imsSumV_pre[q] = (imsSumV_pre[q] || 0) + v_new;
    });

    return inputRows.map(row => {
      const g = (row["G"] as CellInfo)?.value as number || 1;
      const h = (row["H"] as CellInfo)?.value as number || 1;
      const l = (row["L"] as CellInfo)?.value as number || 0;
      const m = (row["M"] as CellInfo)?.value as number || 0;
      const n = (row["N"] as CellInfo)?.value as number || 0;
      const ak = (row["AK"] as CellInfo)?.value as number || 0;
      const ah = (row["AH"] as CellInfo)?.value as number || 0;
      const aj = ak !== 0 ? 1 : 0;
      const delta = sumAK > 0 ? (ak / sumAK - ah / sumAH) : 0;
      const q = String((row["Q"] as CellInfo)?.value || "");

      const t = (l + delta * sumL) * aj;
      const u = (n + delta * sumN) * aj;
      const v2 = (m + delta * sumM) * aj;
      const o = g > 0 ? l / g : 0;
      const p = h > 0 ? n / h : 0;
      const w2 = g > 0 ? t / g : 0;
      const x = h > 0 ? u / h : 0;

      const rVal = imsSumM[q] ?? m;      // R = ΣM same IMS group (Tải IMS Trước cắt)
      const sVal = imsSumV_pre[q] ?? v2; // S = ΣV same IMS group (Tải IMS Sau cắt)

      // IMS License for this row's IMS group
      const imsLic = (row as any)._imsLic as number || 0;
      // Recalculate IMS license based on IMS group (find from tableInputData)
      const o2Val = imsLic > 0 ? rVal / imsLic : 0;
      const o3Val = imsLic > 0 ? sVal / imsLic : 0;

      return {
        ...row,
        O: { ...(row["O"] as CellInfo), value: o },
        P: { ...(row["P"] as CellInfo), value: p },
        R: { ...(row["R"] as CellInfo), value: rVal, formula: "=SUMIF(Q:Q,Q,M:M)", is_formula: true },
        O2: { ...(row["O2"] as CellInfo), value: o2Val, formula: "=R/IMS_LIC", is_formula: true },
        S: { ...(row["S"] as CellInfo), value: sVal, formula: "=SUMIF(Q:Q,Q,V:V)", is_formula: true },
        O3: { ...(row["O3"] as CellInfo), value: o3Val, formula: "=S/IMS_LIC", is_formula: true },
        T: { ...(row["T"] as CellInfo), value: t },
        U: { ...(row["U"] as CellInfo), value: u },
        V: { ...(row["V"] as CellInfo), value: v2 },
        W: { ...(row["W"] as CellInfo), value: w2 },
        X: { ...(row["X"] as CellInfo), value: x },
        AJ: { ...(row["AJ"] as CellInfo), value: aj },
        AL: { ...(row["AL"] as CellInfo), value: ah === ak ? "NO" : "YES" },
        AN: { ...(row["AN"] as CellInfo), value: u - n },
      };
    });
  }

  // ── Cell edit ──────────────────────────────────────────────────────────────
  function handleCellClick(rowIdx: number, colLetter: string, cell: CellInfo) {
    // Find actual index in rows array (not displayRows)
    const dispRow = displayRows[rowIdx];
    const actualIdx = rows.findIndex(r => r.row_num === dispRow.row_num);
    setEditingCell({ rowIdx: actualIdx, colLetter });
    setTempFormula(cell.formula || "");
    setTempValue(String(cell.value ?? ""));
  }

  function handleCellSave() {
    if (!editingCell) return;
    const { rowIdx, colLetter } = editingCell;
    const updatedRows = [...rows];
    const cell = updatedRows[rowIdx][colLetter] as CellInfo;
    if (cell) {
      if (tempFormula.startsWith("=")) {
        cell.is_formula = true;
        cell.formula = tempFormula;
      } else {
        cell.is_formula = false;
        cell.formula = "";
        cell.value = isNaN(Number(tempValue)) ? tempValue : Number(tempValue);
      }
    }
    setRows(recalcRows(updatedRows));
    setEditingCell(null);
  }

  // ── Header rename ──────────────────────────────────────────────────────────
  function saveHeader(col: string) {
    setHeaders(h => ({ ...h, [col]: tempHeaderVal }));
    setEditingHeader(null);
  }

  // ── DB query per table ─────────────────────────────────────────────────────
  async function runDbQuery(tableKey: string) {
    try {
      const res = await axios.post(`${API_BASE}/query-db`, { connection_string: connStr, sql_query: dbQueryMap[tableKey] });
      setDbResults(p => ({ ...p, [tableKey]: res.data }));
      setDbStatuses(p => ({ ...p, [tableKey]: res.data.success ? "success" : "error" }));
      if (res.data.success) applyDbResult(tableKey, res.data.rows);
    } catch { setDbStatuses(p => ({ ...p, [tableKey]: "error" })); }
  }

  function applyDbResult(tableKey: string, dbRows: any[]) {
    if (!dbRows?.length) return;
    setTableInputData(p => {
      const updated = { ...p, [tableKey]: dbRows };
      rebuildSimulationGrid(updated);
      saveTableToBackend(tableKey, dbRows);
      return updated;
    });
  }

  // ── Dashboard data ─────────────────────────────────────────────────────────
  const allAreas = useMemo(() =>
    [...new Set(rows.map(r => String((r["D"] as CellInfo)?.value || "")).filter(Boolean))].sort(),
    [rows]
  );

  const allHwSites = useMemo(() =>
    [...new Set(rows.map(r => String((r["AM"] as CellInfo)?.value || "")).filter(Boolean))].sort(),
    [rows]
  );

  const filteredDashRows = useMemo(() => {
    let result = displayRows;
    if (selectedAreas.length > 0)
      result = result.filter(r => selectedAreas.includes(String((r["D"] as CellInfo)?.value || "")));
    if (selectedHwSites.length > 0)
      result = result.filter(r => selectedHwSites.includes(String((r["AM"] as CellInfo)?.value || "")));
    if (selectedNodes.length > 0)
      result = result.filter(r => selectedNodes.includes(String((r["B"] as CellInfo)?.value || "")));
    return result;
  }, [displayRows, selectedNodes, selectedAreas, selectedHwSites]);

  function buildChart(rows: GridRow[], colPairs: { key: string; col: string; color: string }[]) {
    return rows.map(r => {
      const name = String((r["B"] as CellInfo)?.value || r.row_num);
      const entry: any = { name };
      colPairs.forEach(p => {
        const cell = r[p.col] as CellInfo;
        entry[p.key] = cell?.value ?? 0;
      });
      return entry;
    });
  }

  const COLORS = {
    bear: "#3b82f6", licBear: "#10b981", ipv4: "#8b5cf6",
    bearIMS: "#ec4899", ipv4IMS: "#f59e0b",
    thru: "#06b6d4", licThru: "#84cc16",
    bearDK: "#60a5fa", imsdk: "#f472b6", thruDK: "#22d3ee",
  };

  // Merged chart data: combine before + after columns per node in single dataset
  const bearChartData = filteredDashRows.map(r => {
    const name = String((r["B"] as CellInfo)?.value || r.row_num);
    return {
      name,
      "Bear Total sử dụng": (r["L"] as CellInfo)?.value ?? 0,
      "Bear sử dụng DK": (r["T"] as CellInfo)?.value ?? 0,
      "License Bear UCTT": (r["G"] as CellInfo)?.value ?? 0,
      "Total IPv4 (v-internet)": (r["I"] as CellInfo)?.value ?? 0,
    };
  });

  const imsChartData = filteredDashRows.map(r => {
    const name = String((r["B"] as CellInfo)?.value || r.row_num);
    return {
      name,
      "Bear IMS": (r["M"] as CellInfo)?.value ?? 0,
      "Bearer IMS dự kiến": (r["V"] as CellInfo)?.value ?? 0,
      "Total IPv4 (IMS)": (r["I2"] as CellInfo)?.value ?? 0,
    };
  });

  const thruChartData = filteredDashRows.map(r => {
    const name = String((r["B"] as CellInfo)?.value || r.row_num);
    return {
      name,
      "Throughput": (r["N"] as CellInfo)?.value ?? 0,
      "Throughput dự kiến": (r["U"] as CellInfo)?.value ?? 0,
      "License Throughput VHKT": (r["H"] as CellInfo)?.value ?? 0,
    };
  });

  // Chart 4: % Tải IMS / License - group by unique IMS systems (Q values)
  const imsPctChartData = useMemo(() => {
    // Group filtered rows by IMS system (Q), take the first O2/O3 per group (they're same for same IMS group)
    const imsGroups: Record<string, { o2: number; o3: number }> = {};
    filteredDashRows.forEach(r => {
      const imsGroup = String((r["Q"] as CellInfo)?.value || "");
      if (!imsGroup) return;
      if (!imsGroups[imsGroup]) {
        imsGroups[imsGroup] = {
          o2: Number((r["O2"] as CellInfo)?.value || 0),
          o3: Number((r["O3"] as CellInfo)?.value || 0),
        };
      }
    });
    return Object.entries(imsGroups).map(([name, vals]) => ({
      name,
      "% Tải IMS trước": vals.o2,
      "% Tải IMS sau": vals.o3,
    }));
  }, [filteredDashRows]);

  const tooltipStyle = { backgroundColor: "#0f172a", borderColor: "#1e293b", borderRadius: 12 };

  // ── Render helpers ─────────────────────────────────────────────────────────
  function ChartCard({
    title, data, bars,
  }: {
    title: string;
    data: any[];
    bars: { key: string; color: string; type?: "before" | "after" | "ref" }[];
  }) {
    return (
      <div className="glass-card rounded-3xl p-5">
        <div className="flex items-start justify-between mb-3 gap-4">
          <h4 className="text-sm font-bold text-gray-300">{title}</h4>
          <div className="flex items-center gap-3 shrink-0 text-[10px] text-gray-500">
            <span className="flex items-center gap-1">
              <span className="inline-block w-3 h-2 rounded-sm opacity-100" style={{ background: "#3b82f6" }} />
              Trước cắt
            </span>
            <span className="flex items-center gap-1">
              <span className="inline-block w-3 h-2 rounded-sm opacity-80 border border-white/20" style={{ background: "#60a5fa" }} />
              Sau cắt
            </span>
            <span className="flex items-center gap-1">
              <span className="inline-block w-3 h-2 rounded-sm" style={{ background: "#10b981" }} />
              Tham chiếu
            </span>
          </div>
        </div>
        <div className="h-[420px]">
          <ResponsiveContainer width="100%" height="100%">
            <BarChart data={data} margin={{ top: 5, right: 20, left: 10, bottom: 70 }}>
              <CartesianGrid strokeDasharray="3 3" stroke="#1e293b" />
              <XAxis dataKey="name" stroke="#6b7280" tick={{ fontSize: 11 }} angle={-35} textAnchor="end" interval={0} />
              <YAxis stroke="#6b7280" tick={{ fontSize: 10 }} tickFormatter={v => v >= 1e6 ? `${(v / 1e6).toFixed(1)}M` : v >= 1e3 ? `${(v / 1e3).toFixed(0)}K` : String(v)} />
              <Tooltip
                contentStyle={tooltipStyle}
                formatter={(value: any, name: any) => [
                  value >= 1e6 ? `${(value / 1e6).toFixed(2)}M` : value >= 1e3 ? `${(value / 1e3).toFixed(1)}K` : Number(value || 0).toLocaleString(),
                  name
                ]}
              />
              <Legend wrapperStyle={{ paddingTop: "14px", fontSize: "11px" }} />
              {bars.map(b => (
                <Bar
                  key={b.key}
                  dataKey={b.key}
                  fill={b.color}
                  radius={[3, 3, 0, 0]}
                  fillOpacity={b.type === "after" ? 0.7 : b.type === "ref" ? 0.5 : 1}
                  stroke={b.type === "after" ? b.color : "transparent"}
                  strokeWidth={b.type === "after" ? 1.5 : 0}
                  strokeDasharray={b.type === "after" ? "4 2" : "0"}
                />
              ))}
            </BarChart>
          </ResponsiveContainer>
        </div>
      </div>
    );
  }

  function TableInputPanel({ tableKey }: { tableKey: string }) {
    const tmpl = TABLE_TEMPLATES[tableKey as keyof typeof TABLE_TEMPLATES];
    const data = tableInputData[tableKey] || [];
    const status = dbStatuses[tableKey];
    const [xlStatus, setXlStatus] = React.useState<"idle" | "ok" | "err">("idle");
    const [xlMsg, setXlMsg] = React.useState("");

    const page = tablePages[tableKey] || 1;
    const startIndex = (page - 1) * tablePageSize;
    const pageData = data.slice(startIndex, startIndex + tablePageSize);
    const totalPages = Math.ceil(data.length / tablePageSize) || 1;

    async function handleTableExcel(e: React.ChangeEvent<HTMLInputElement>) {
      const file = e.target.files?.[0];
      if (!file) return;
      setXlStatus("idle");
      setXlMsg("");
      try {
        const buf = await file.arrayBuffer();
        const wb = XLSX.read(buf, { type: "array" });
        const sheetName = wb.SheetNames[0];
        const ws = wb.Sheets[sheetName];
        const rawJson: any[] = XLSX.utils.sheet_to_json(ws, { defval: "" });
        if (rawJson.length === 0) {
          setXlStatus("err");
          setXlMsg("File không có dữ liệu.");
        } else {
          // Validate structure & predefined columns
          const firstRow = rawJson[0];
          const rowKeys = Object.keys(firstRow).map(k => k.trim().toLowerCase());
          const missingCols = tmpl.columns.filter(col => {
            const colLower = col.trim().toLowerCase();
            return !rowKeys.some(rk => rk.includes(colLower) || colLower.includes(rk));
          });
          
          if (missingCols.length > 0) {
            setXlStatus("err");
            setXlMsg(`File import sai cấu trúc. Thiếu cột: ${missingCols.join(", ")}`);
            e.target.value = "";
            return;
          }

          // Normalize columns, ensure Vendor exists
          const normalized = rawJson.map(row => {
            const nodeKey = Object.keys(row).find(k => k.toLowerCase() === "node");
            const vendorKey = Object.keys(row).find(k => k.toLowerCase() === "vendor");
            return {
              ...row,
              Node: nodeKey ? row[nodeKey] : (row.Node || ""),
              Vendor: vendorKey ? row[vendorKey] : (row.Vendor || "Huawei")
            };
          });

          setTableInputData(p => {
            const updated = { ...p, [tableKey]: normalized };
            rebuildSimulationGrid(updated);
            saveTableToBackend(tableKey, normalized);
            return updated;
          });
          setTablePages(p => ({ ...p, [tableKey]: 1 }));
          setXlStatus("ok");
          setXlMsg(`Đã đọc ${normalized.length} dòng từ sheet "${sheetName}"`);
        }
      } catch (err: any) {
        setXlStatus("err");
        setXlMsg("Không đọc được file. Kiểm tra định dạng .xlsx.");
      }
      e.target.value = "";
    }

    return (
      <div className="glass-card rounded-3xl p-5 space-y-4">
        <div className="flex items-start justify-between">
          <div>
            <h4 className="font-bold text-sm text-gray-200">{tmpl.label}</h4>
            <p className="text-xs text-gray-500 mt-0.5">{tmpl.description}</p>
          </div>
          <span className="text-[10px] bg-gray-800 text-gray-400 px-2 py-0.5 rounded-full">{tmpl.columns.length} cột</span>
        </div>

        {/* Excel upload for this table */}
        <div className={`relative border border-dashed rounded-2xl p-4 flex flex-col gap-2 transition group cursor-pointer
          ${xlStatus === "ok" ? "border-emerald-600/60 bg-emerald-950/10" : xlStatus === "err" ? "border-red-500/40 bg-red-950/10" : "border-gray-700 hover:border-blue-500/60"}`}>
          <input type="file" accept=".xlsx,.xls" onChange={handleTableExcel}
            className="absolute inset-0 opacity-0 cursor-pointer z-10" />
          <div className="flex items-center gap-3">
            <div className={`p-2.5 rounded-xl transition ${xlStatus === "ok" ? "bg-emerald-600/20" : xlStatus === "err" ? "bg-red-500/10" : "bg-gray-800 group-hover:bg-blue-600/20"}`}>
              {xlStatus === "ok"
                ? <CheckCircle2 className="w-5 h-5 text-emerald-400" />
                : xlStatus === "err"
                  ? <AlertTriangle className="w-5 h-5 text-red-400" />
                  : <UploadCloud className="w-5 h-5 text-gray-400 group-hover:text-blue-400 transition" />}
            </div>
            <div className="flex-1 min-w-0">
              <p className="text-xs font-semibold text-gray-300">Import từ Excel (.xlsx)</p>
              <p className="text-[10px] text-gray-500">Hàng đầu tiên trong file = tên cột. Click hoặc kéo thả.</p>
            </div>
          </div>
          {xlMsg && (
            <p className={`text-[10px] font-medium pl-1 ${xlStatus === "ok" ? "text-emerald-400" : "text-red-400"}`}>{xlMsg}</p>
          )}
        </div>

        {/* Inline table editor — columns from imported file OR template */}
        {(() => {
          const importedCols = data.length > 0
            ? Object.keys(data[0])
            : tmpl.columns;
          const displayCols = importedCols.length > 0 ? importedCols : tmpl.columns;
          return (
            <div className="space-y-3">
              <div className="overflow-x-auto border border-gray-800/60 rounded-2xl">
                <table className="w-full text-xs text-left border-collapse">
                  <thead>
                    <tr className="bg-gray-900/60 border-b border-gray-800">
                      {displayCols.map(col => (
                        <th key={col} className="px-3 py-2 font-bold text-gray-400 whitespace-nowrap">{col}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {pageData.map((row, ri) => (
                      <tr key={ri} className="border-b border-gray-800/40 hover:bg-gray-800/10">
                        {displayCols.map(col => (
                          <td key={col} className="px-1 py-1">
                            <input
                              className="w-full bg-black/40 border border-gray-800 rounded px-2 py-1 text-gray-300 focus:outline-none focus:border-blue-500 min-w-[80px]"
                              value={row[col] ?? ""}
                              onChange={e => {
                                const globalIndex = startIndex + ri;
                                const updatedData = data.map((r, i) => i === globalIndex ? { ...r, [col]: e.target.value } : r);
                                setTableInputData(p => {
                                  const nextState = { ...p, [tableKey]: updatedData };
                                  rebuildSimulationGrid(nextState);
                                  saveTableToBackend(tableKey, updatedData);
                                  return nextState;
                                });
                              }}
                            />
                          </td>
                        ))}
                      </tr>
                    ))}
                    {pageData.length === 0 && (
                      <tr>
                        <td colSpan={displayCols.length} className="text-center py-6 text-gray-500">
                          Chưa có dữ liệu. Hãy import file Excel hoặc chạy SQL.
                        </td>
                      </tr>
                    )}
                  </tbody>
                </table>
              </div>

              {/* Pagination controls */}
              {data.length > 0 && (
                <div className="flex items-center justify-between text-xs px-2 py-1 bg-gray-900/30 rounded-xl border border-gray-800/60">
                  <div className="text-gray-400">
                    Hiển thị {startIndex + 1} - {Math.min(startIndex + tablePageSize, data.length)} / {data.length} dòng
                  </div>
                  <div className="flex items-center gap-2">
                    <button
                      disabled={page === 1}
                      onClick={() => setTablePages(p => ({ ...p, [tableKey]: page - 1 }))}
                      className="px-2.5 py-1 bg-gray-800 hover:bg-gray-700 disabled:opacity-40 disabled:hover:bg-gray-800 rounded text-gray-300 transition"
                    >
                      Trước
                    </button>
                    <span className="text-gray-300 font-medium">Trang {page} / {totalPages}</span>
                    <button
                      disabled={startIndex + tablePageSize >= data.length}
                      onClick={() => setTablePages(p => ({ ...p, [tableKey]: page + 1 }))}
                      className="px-2.5 py-1 bg-gray-800 hover:bg-gray-700 disabled:opacity-40 disabled:hover:bg-gray-800 rounded text-gray-300 transition"
                    >
                      Sau
                    </button>
                  </div>
                </div>
              )}
            </div>
          );
        })()}

        {/* SQL query section */}
        <div className="space-y-3 border-t border-gray-800 pt-4">
          <label className="text-[10px] text-gray-500 font-bold uppercase tracking-wider flex items-center gap-1.5">
            <Database className="w-3 h-3" /> Hoặc nhập từ SQL
          </label>
          <textarea
            rows={2}
            value={dbQueryMap[tableKey] || ""}
            onChange={e => setDbQueryMap(p => ({ ...p, [tableKey]: e.target.value }))}
            className="w-full bg-black/40 border border-gray-800 rounded-xl px-3 py-2 text-xs font-mono text-gray-300 focus:outline-none focus:border-blue-500 resize-none"
          />
          
          <div className="flex flex-wrap items-center justify-between gap-3 bg-gray-900/25 p-3 rounded-2xl border border-gray-800/40">
            <div className="flex flex-wrap items-center gap-3">
              <button
                onClick={() => runDbQuery(tableKey)}
                className="px-3 py-1.5 bg-indigo-600 hover:bg-indigo-500 text-white rounded-xl text-xs font-bold flex items-center gap-1.5 transition"
              >
                <Play className="w-3 h-3" /> Chạy SQL Ngay
              </button>
              
              <div className="flex items-center gap-2">
                <span className="text-[10px] text-gray-400 font-semibold">Chu kỳ lập lịch:</span>
                <select
                  value={scheduleTypes[tableKey] || "manual"}
                  onChange={e => setScheduleTypes(p => ({ ...p, [tableKey]: e.target.value }))}
                  className="bg-black/60 border border-gray-800 rounded-lg px-2 py-1 text-xs text-gray-300 focus:outline-none focus:border-blue-500"
                >
                  <option value="manual">Chạy tay (Manual)</option>
                  <option value="5m">Mỗi 5 phút</option>
                  <option value="30m">Mỗi 30 phút</option>
                  <option value="1h">Mỗi 1 giờ</option>
                  <option value="24h">Hàng ngày (24h)</option>
                </select>
              </div>

              <div className="flex items-center gap-1.5">
                <input
                  type="checkbox"
                  id={`sched-active-${tableKey}`}
                  checked={scheduleActives[tableKey] || false}
                  onChange={e => setScheduleActives(p => ({ ...p, [tableKey]: e.target.checked }))}
                  className="rounded border-gray-800 text-blue-600 focus:ring-blue-500 bg-black/40"
                />
                <label htmlFor={`sched-active-${tableKey}`} className="text-[10px] text-gray-300 font-semibold select-none cursor-pointer">Kích hoạt</label>
              </div>
            </div>

            <button
              onClick={() => saveSchedule(tableKey)}
              className="px-3 py-1.5 bg-gray-800 hover:bg-gray-700 text-gray-300 rounded-xl text-xs font-semibold border border-gray-700/60 transition"
            >
              Lưu Lập Lịch
            </button>
          </div>

          {schedules.find(s => s.table_key === tableKey)?.last_run && (
            <p className="text-[10px] text-gray-500">
              Lần chạy cuối: {new Date(schedules.find(s => s.table_key === tableKey).last_run).toLocaleString("vi-VN")}
            </p>
          )}

          <div className="flex items-center gap-3">
            {status === "success" && <span className="text-xs text-emerald-400 flex items-center gap-1"><CheckCircle2 className="w-3 h-3" /> Chạy thành công</span>}
            {status === "error" && <span className="text-xs text-red-400 flex items-center gap-1"><AlertTriangle className="w-3 h-3" /> Lỗi chạy SQL</span>}
          </div>
        </div>
      </div>
    );
  }

  // ── Visible columns (exclude I2 from separate rendering — handled inline) ──
  const visibleCols = Object.keys(headers).filter(k => k !== "I2");

  const pctColOverloads = displayRows.some(r =>
    ["O", "P", "W", "X"].some(c => ((r[c] as CellInfo)?.value || 0) * 100 > overloadThreshold)
  );

  return (
    <div className="min-h-screen flex bg-[#080b11] text-[#f3f4f6]">
      {/* ─ Sidebar ──────────────────────────────────────────────────────────── */}
      <aside className="w-64 glass border-r border-[#1f293d] flex flex-col justify-between p-6 shrink-0">
        <div>
          <div className="flex items-center gap-3 mb-8">
            <div className="p-2.5 bg-gradient-to-br from-blue-600 to-indigo-600 rounded-xl shadow-lg shadow-blue-500/20">
              <FileSpreadsheet className="w-6 h-6 text-white" />
            </div>
            <div>
              <h1 className="font-bold text-base leading-tight bg-gradient-to-r from-blue-400 to-indigo-400 bg-clip-text text-transparent">GGSN Optimizer</h1>
            </div>
          </div>

          <nav className="space-y-1">
            <button
              onClick={() => setActiveTab("simulation")}
              className={`w-full flex items-center gap-3 px-4 py-3 rounded-xl text-sm font-medium transition-all duration-200 ${activeTab === "simulation"
                ? "bg-blue-600/10 text-blue-400 border border-blue-500/20"
                : "text-gray-400 hover:text-gray-200 hover:bg-gray-800/40"}`}
            >
              <Table2 className="w-4 h-4" />
              Bảng Mô Phỏng Tải
            </button>

            <div>
              <button
                onClick={() => {
                  setActiveTab("sources");
                  setSourcesExpanded(!sourcesExpanded);
                }}
                className={`w-full flex items-center justify-between px-4 py-3 rounded-xl text-sm font-medium transition-all duration-200 ${activeTab === "sources"
                  ? "bg-blue-600/5 text-blue-400/90"
                  : "text-gray-400 hover:text-gray-200 hover:bg-gray-800/40"}`}
              >
                <div className="flex items-center gap-3">
                  <Database className="w-4 h-4" />
                  <span>Nguồn Dữ Liệu</span>
                </div>
                {sourcesExpanded ? <ChevronUp className="w-3.5 h-3.5" /> : <ChevronDown className="w-3.5 h-3.5" />}
              </button>

              {sourcesExpanded && (
                <div className="mt-1 ml-4 pl-2 border-l border-gray-800 space-y-1">
                  {[
                    { key: "license", label: "Bảng License GGSN" },
                    { key: "current", label: "Bảng Tải Hiện Tại" },
                    { key: "weight", label: "Bảng Trọng Số Cân Tải" },
                    { key: "ims_routing", label: "Bảng IMS Routing" },
                    { key: "hw_site", label: "Bảng HW/NFVI Site" },
                    { key: "ims_license", label: "Bảng License IMS" },
                  ].map(sub => (
                    <button
                      key={sub.key}
                      onClick={() => {
                        setActiveTab("sources");
                        setCurrentSourceTab(sub.key as any);
                      }}
                      className={`w-full text-left px-3 py-2 rounded-lg text-xs font-medium transition-all duration-150 ${activeTab === "sources" && currentSourceTab === sub.key
                        ? "bg-blue-600/10 text-blue-400 font-semibold"
                        : "text-gray-500 hover:text-gray-300 hover:bg-gray-800/30"}`}
                    >
                      • {sub.label}
                    </button>
                  ))}
                </div>
              )}
            </div>

            <button
              onClick={() => setActiveTab("dashboard")}
              className={`w-full flex items-center gap-3 px-4 py-3 rounded-xl text-sm font-medium transition-all duration-200 ${activeTab === "dashboard"
                ? "bg-blue-600/10 text-blue-400 border border-blue-500/20"
                : "text-gray-400 hover:text-gray-200 hover:bg-gray-800/40"}`}
            >
              <BarChart3 className="w-4 h-4" />
              Biểu Đồ Tương Quan
            </button>
          </nav>
        </div>

        <div className="space-y-4">
          <div className="p-4 bg-gray-900/40 rounded-xl border border-gray-800/50 space-y-3">
            <div className="flex items-center gap-2 text-[10px] font-bold text-gray-400 uppercase tracking-wider">
              <SlidersHorizontal className="w-3.5 h-3.5" /> Cấu hình chung
            </div>
            <div>
              <label className="block text-[10px] text-gray-500 mb-1">Ngưỡng quá tải (%):</label>
              <input type="number" value={overloadThreshold} onChange={e => setOverloadThreshold(Number(e.target.value))}
                className="w-full bg-black/40 border border-gray-800 rounded px-2.5 py-1.5 text-sm text-gray-300 focus:outline-none focus:border-blue-500" />
            </div>
            <div className="flex items-center gap-2 mt-2">
              <span className="w-2 h-2 bg-red-500 rounded-sm"></span>
              <span className="text-[10px] text-gray-500">% vượt ngưỡng</span>
            </div>
            <div className="flex items-center gap-2">
              <span className="w-2 h-2 bg-amber-500 rounded-sm"></span>
              <span className="text-[10px] text-gray-500">% Top 5 cao nhất</span>
            </div>
            <div className="flex items-center gap-2">
              <span className="w-2 h-2 bg-blue-500 rounded-sm"></span>
              <span className="text-[10px] text-gray-500">Số Top 5 cao nhất</span>
            </div>
          </div>
        </div>
      </aside>

      {/* ─ Main ─────────────────────────────────────────────────────────────── */}
      <main className="flex-1 overflow-auto flex flex-col min-w-0">
        {/* Header */}
        <header className="h-14 border-b border-[#1f293d] glass flex items-center justify-between px-6 z-10 shrink-0">
          <div className="flex items-center gap-3">
            <div className="text-xs px-2.5 py-1 bg-gray-800 text-gray-300 rounded-full border border-gray-700/50 flex items-center gap-1.5 font-medium">
              <span className="w-1.5 h-1.5 rounded-full bg-emerald-500 animate-pulse"></span>
              {fileName}
            </div>
          </div>
          <div className="flex items-center gap-2">
            <button onClick={fetchDbData} className="px-3 py-1.5 bg-gray-800 hover:bg-gray-700 text-gray-300 rounded-xl text-xs font-semibold flex items-center gap-1.5 border border-gray-700/60 transition">
              <RefreshCw className="w-3.5 h-3.5" /> Reset
            </button>
            <button onClick={() => {
              try {
                const wb = XLSX.utils.book_new();

                // Sheet 1: Bảng Mô Phỏng Tải
                const simCols = Object.keys(headers);
                const simHeader: string[] = simCols.map(k => headers[k] || k);
                const simData: (string | number)[][] = [simHeader];
                rows.forEach(row => {
                  simData.push(simCols.map(k => {
                    const cell = row[k] as CellInfo;
                    const v = cell?.value;
                    return (typeof v === "number" || typeof v === "string") ? v : "";
                  }));
                });
                const simWs = XLSX.utils.aoa_to_sheet(simData);
                XLSX.utils.book_append_sheet(wb, simWs, "Mô Phỏng Tải");

                // Sheet 2-4: Các bảng nguồn dữ liệu
                const tableLabels: Record<string, string> = {
                  license: "License Node",
                  current: "Tải Hiện Tại",
                  weight: "Trọng Số Cân Tải",
                };
                Object.entries(tableInputData).forEach(([key, tableRows]) => {
                  if (!tableRows || tableRows.length === 0) return;
                  const tCols = Object.keys(tableRows[0]);
                  const tData: (string | number)[][] = [tCols];
                  tableRows.forEach((r: Record<string, unknown>) => {
                    tData.push(tCols.map(c => {
                      const v = r[c];
                      return (typeof v === "number" || typeof v === "string") ? v : "";
                    }));
                  });
                  const ws = XLSX.utils.aoa_to_sheet(tData);
                  XLSX.utils.book_append_sheet(wb, ws, tableLabels[key] || key);
                });

                XLSX.writeFile(wb, `Export_GGSN_${new Date().toISOString().slice(0, 10)}.xlsx`);
              } catch (err) {
                alert("Export thất bại: " + String(err));
              }
            }} className="px-3 py-1.5 bg-gradient-to-r from-blue-600 to-indigo-600 hover:from-blue-500 hover:to-indigo-500 text-white rounded-xl text-xs font-semibold flex items-center gap-1.5 shadow shadow-blue-500/10 transition">
              <Download className="w-3.5 h-3.5" /> Xuất Excel
            </button>
          </div>
        </header>

        <div className="flex-1 p-6 overflow-auto">
          {error && (
            <div className="mb-4 p-3 bg-red-900/20 border border-red-500/20 text-red-400 rounded-2xl flex items-center gap-2 text-sm">
              <AlertCircle className="w-4 h-4 shrink-0" /> {error}
            </div>
          )}

          {/* ══ TAB: NGUỒN DỮ LIỆU ════════════════════════════════════════════ */}
          {activeTab === "sources" && (
            <div className="space-y-5 max-w-6xl">
              {/* Header with name of the active table */}
              <div className="flex items-center justify-between border-b border-gray-800 pb-3">
                <div>
                  <h2 className="text-xl font-bold text-gray-200">
                    Cấu hình nguồn dữ liệu: {TABLE_TEMPLATES[currentSourceTab]?.label}
                  </h2>
                  <p className="text-xs text-gray-500 mt-1">
                    Nhập dữ liệu bằng file Excel hoặc cấu hình DB kết nối trực tiếp.
                  </p>
                </div>
              </div>

              <div className="space-y-5">
                <div className="p-4 glass-card rounded-2xl flex items-center gap-3">
                  <Database className="w-4 h-4 text-indigo-400 shrink-0" />
                  <div className="flex-1">
                    <label className="text-xs text-gray-400 font-semibold">Connection String chung (dùng cho tất cả bảng SQL)</label>
                    <input value={connStr} onChange={e => setConnStr(e.target.value)}
                      className="w-full mt-1 bg-black/40 border border-gray-800 rounded-xl px-3 py-2 text-sm text-gray-300 focus:outline-none focus:border-blue-500"
                      placeholder="sqlite://mock  hoặc  postgresql://..." />
                  </div>
                </div>

                <div className="grid grid-cols-1 gap-5">
                  <TableInputPanel tableKey={currentSourceTab} />
                </div>
              </div>
            </div>
          )}

          {/* ══ TAB: SIMULATION GRID ══════════════════════════════════════════ */}
          {activeTab === "simulation" && (
            <div className="space-y-4">
              {/* Cell editor bar */}
              {editingCell && (
                <div className="p-4 bg-blue-600/10 border border-blue-500/20 rounded-2xl flex flex-wrap items-center justify-between gap-4">
                  <div className="flex flex-wrap items-center gap-4 flex-1">
                    <span className="text-xs font-bold text-blue-400 bg-blue-500/10 px-2.5 py-1 rounded-lg shrink-0">
                      Dòng {rows[editingCell.rowIdx]?.row_num}, Cột {editingCell.colLetter}
                    </span>
                    <div className="flex gap-3 flex-1 min-w-[300px]">
                      <div className="flex-1">
                        <label className="text-[10px] text-gray-500 font-bold uppercase block mb-0.5">Giá Trị Tĩnh</label>
                        <input value={tempValue} onChange={e => setTempValue(e.target.value)}
                          className="w-full bg-black/60 border border-gray-800 rounded-lg px-3 py-1.5 text-sm focus:outline-none" placeholder="VD: 100000" />
                      </div>
                      <div className="flex-[2]">
                        <label className="text-[10px] text-gray-500 font-bold uppercase block mb-0.5">Công Thức (bắt đầu với =)</label>
                        <input value={tempFormula} onChange={e => setTempFormula(e.target.value)}
                          className="w-full bg-black/60 border border-gray-800 rounded-lg px-3 py-1.5 text-sm font-mono focus:outline-none focus:border-blue-500" placeholder="=L3*50%" />
                      </div>
                    </div>
                  </div>
                  <div className="flex gap-2">
                    <button onClick={handleCellSave} className="px-4 py-2 bg-blue-600 hover:bg-blue-500 text-white rounded-xl text-xs font-bold flex items-center gap-1.5 transition">
                      <Check className="w-3.5 h-3.5" /> Lưu
                    </button>
                    <button onClick={() => setEditingCell(null)} className="px-3 py-2 bg-gray-800 hover:bg-gray-700 text-gray-400 rounded-xl text-xs font-semibold transition">
                      <X className="w-3.5 h-3.5" />
                    </button>
                  </div>
                </div>
              )}

              {/* Grid */}
              {isFullView && (
                <div
                  className="fixed inset-0 z-40 bg-black/60 backdrop-blur-sm"
                  onClick={() => setIsFullView(false)}
                />
              )}
              <div
                onKeyDown={e => { if (e.key === "Escape") setIsFullView(false); }}
                tabIndex={-1}
                className={`glass-card rounded-3xl p-5 overflow-hidden outline-none transition-all duration-300
                  ${isFullView
                    ? "fixed inset-4 z-50 shadow-2xl shadow-black/80 border border-blue-700/40"
                    : ""}
                `}>
                <div className="flex items-center justify-between mb-4">
                  <div>
                    <h3 className="font-bold text-base">Bảng Tính Toán Cân Tải GGSN</h3>
                    <p className="text-[11px] text-gray-500">
                      Click ô để chỉnh sửa • <span className="text-emerald-400">Xanh = công thức</span> •
                      <span className="text-red-400 ml-1">Đỏ = vượt ngưỡng</span> •
                      <span className="text-amber-400 ml-1">Vàng = top 5%</span> •
                      <span className="text-blue-400 ml-1">Xanh dương = top 5 số</span>
                    </p>
                  </div>
                  <div className="flex items-center gap-3">
                    <div className="flex items-center gap-1.5 bg-gray-900/60 p-1.5 rounded-xl border border-gray-800/60 text-[11px] text-gray-400">
                      <span className="font-semibold px-1">Thu Phóng:</span>
                      <button
                        onClick={() => setZoomScale(s => Math.max(75, s - 5))}
                        className="px-2 py-0.5 bg-gray-800 hover:bg-gray-700 text-gray-300 rounded font-bold"
                      >
                        -
                      </button>
                      <span className="font-mono text-gray-200 min-w-[35px] text-center">{zoomScale}%</span>
                      <button
                        onClick={() => setZoomScale(s => Math.min(150, s + 5))}
                        className="px-2 py-0.5 bg-gray-800 hover:bg-gray-700 text-gray-300 rounded font-bold"
                      >
                        +
                      </button>
                      <button
                        onClick={() => setZoomScale(100)}
                        className="px-1.5 py-0.5 bg-gray-800 hover:bg-gray-700 text-gray-400 hover:text-gray-200 rounded"
                      >
                        Reset
                      </button>
                    </div>

                    <button onClick={() => setSortAsc(v => !v)}
                      className="flex items-center gap-1.5 px-3 py-1.5 text-xs bg-gray-800 hover:bg-gray-700 text-gray-300 rounded-xl border border-gray-700 transition">
                      {sortAsc ? <ChevronUp className="w-3.5 h-3.5" /> : <ChevronDown className="w-3.5 h-3.5" />}
                      Node {sortAsc ? "A→Z" : "Z→A"}
                    </button>

                    <button
                      onClick={() => setIsFullView(v => !v)}
                      title={isFullView ? "Thu nhỏ" : "Phóng to toàn màn hình"}
                      className="flex items-center gap-1.5 px-3 py-1.5 text-xs bg-blue-600/20 hover:bg-blue-600/40 text-blue-300 hover:text-blue-100 rounded-xl border border-blue-700/50 hover:border-blue-500 transition">
                      {isFullView ? <Minimize2 className="w-3.5 h-3.5" /> : <Maximize2 className="w-3.5 h-3.5" />}
                      {isFullView ? "Thu nhỏ" : "Toàn màn hình"}
                    </button>
                  </div>
                </div>

                <div
                  className={`overflow-auto border border-gray-800/60 rounded-2xl transition-all duration-300 ${isFullView ? "max-h-[calc(100vh-160px)]" : "max-h-[600px]"}`}
                  onClick={e => { if (!(e.ctrlKey || e.metaKey || e.shiftKey)) setSelectedCells(new Set()); }}>
                  <table className="text-left border-collapse text-xs w-max min-w-full" style={{ fontSize: `${zoomScale}%` }}>
                    <thead className="sticky top-0 z-20">
                      {/* Row 1 — group labels */}
                      <tr className="bg-gray-950 border-b border-gray-800/50">
                        <th className="px-3 py-2 text-gray-600 border-r border-gray-800 sticky left-0 z-30 bg-gray-950" colSpan={1}>Dòng</th>
                        {visibleCols.map((col, ci) => {
                          const grp = headersRow1[col] || "";
                          const isSticky = ci < 2;
                          const stickyLeft = ci === 0 ? "left-10" : "left-[calc(2.5rem+110px)]";
                          return (
                            <th key={`g-${col}`}
                              className={`px-2 py-2 text-[10px] font-semibold text-gray-500 text-center border-r border-gray-800 whitespace-nowrap bg-gray-950
                ${isSticky ? `sticky ${stickyLeft} z-30` : ""}`}>
                              {ci < 2 ? (ci === 0 ? "Node" : "Vendor") : grp}
                            </th>
                          );
                        })}
                      </tr>
                      {/* Row 2 — field names */}
                      <tr className="bg-gray-900/70 border-b border-gray-800">
                        <th className="px-3 py-2.5 text-gray-400 font-bold border-r border-gray-800 w-10 text-center sticky left-0 z-30 bg-gray-900/95">Dòng</th>
                        {visibleCols.map((col, ci) => {
                          const isSticky = ci < 2;
                          const stickyLeft = ci === 0 ? "left-10" : "left-[calc(2.5rem+110px)]";
                          return (
                            <th key={col}
                              className={`px-2 py-2.5 border-r border-gray-800 whitespace-nowrap bg-gray-900/95
                ${isSticky ? `sticky ${stickyLeft} z-30` : ""}`}>
                              <div className="text-[10px] text-blue-500/80 font-bold mb-0.5">{col}</div>
                              {editingHeader === col ? (
                                <div className="flex items-center gap-1">
                                  <input autoFocus value={tempHeaderVal} onChange={e => setTempHeaderVal(e.target.value)}
                                    onKeyDown={e => { if (e.key === "Enter") saveHeader(col); if (e.key === "Escape") setEditingHeader(null); }}
                                    className="bg-black/60 border border-blue-500 rounded px-1.5 py-0.5 text-xs w-24 focus:outline-none" />
                                  <button onClick={() => saveHeader(col)} className="text-emerald-400 hover:text-emerald-300"><Check className="w-3 h-3" /></button>
                                  <button onClick={() => setEditingHeader(null)} className="text-gray-500 hover:text-gray-300"><X className="w-3 h-3" /></button>
                                </div>
                              ) : (
                                <div className="flex items-center gap-1 group text-gray-300 font-semibold">
                                  <span className="max-w-[120px] truncate">{headers[col]}</span>
                                  <button onClick={() => { setEditingHeader(col); setTempHeaderVal(headers[col]); }}
                                    className="opacity-0 group-hover:opacity-100 text-gray-600 hover:text-blue-400 transition ml-1">
                                    <Edit2 className="w-2.5 h-2.5" />
                                  </button>
                                </div>
                              )}
                            </th>
                          );
                        })}
                      </tr>
                      {/* Row 3 — column filters */}
                      <tr className="bg-gray-900/40 border-b border-gray-800">
                        <th className="px-2 py-1 border-r border-gray-800 sticky left-0 z-30 bg-gray-950/95">
                          <span className="sr-only">Lọc</span>
                        </th>
                        {visibleCols.map((col, ci) => {
                          const isSticky = ci < 2;
                          const stickyLeft = ci === 0 ? "left-10" : "left-[calc(2.5rem+110px)]";
                          return (
                            <th key={`f-${col}`}
                              className={`px-1 py-1 border-r border-gray-800 bg-gray-900/95
                                ${isSticky ? `sticky ${stickyLeft} z-30` : ""}`}>
                              <input
                                value={columnFilters[col] || ""}
                                onChange={e => setColumnFilters(prev => ({ ...prev, [col]: e.target.value }))}
                                placeholder="Lọc..."
                                className="w-full bg-black/40 border border-gray-800/80 rounded px-1.5 py-0.5 text-[10px] text-gray-300 focus:outline-none focus:border-blue-500/60 min-w-[70px]"
                              />
                            </th>
                          );
                        })}
                      </tr>
                    </thead>
                    <tbody>
                      {displayRows.map((row, rowIdx) => {
                        const expPct = ((row["X"] as CellInfo)?.value || 0) * 100;
                        const isOverloaded = expPct > overloadThreshold;
                        const rowKey = String(row.row_num);
                        const isRowSelected = visibleCols.some(c => selectedCells.has(`${rowKey}:${c}`));
                        const isEditing = (col: string) =>
                          editingCell?.rowIdx === rows.findIndex(r => r.row_num === row.row_num) &&
                          editingCell?.colLetter === col;

                        return (
                          <tr key={rowIdx}
                            className={`border-b border-gray-800/30 transition
              ${isOverloaded ? "bg-red-950/10" : ""}
              ${isRowSelected ? "bg-blue-950/25" : "hover:bg-gray-800/15"}`}>
                            <td className="px-2 py-1.5 text-center text-gray-600 font-bold bg-gray-950/30 border-r border-gray-800 sticky left-0 z-10">{row.row_num}</td>
                            {visibleCols.map((col, ci) => {
                              const cell = row[col] as CellInfo;
                              if (!cell) return <td key={col} className="px-3 py-1.5 border-r border-gray-800/40 text-gray-700">—</td>;

                              const isFormula = cell.is_formula;
                              const val = cell.value;
                              const cellId = `${rowKey}:${col}`;
                              const isCellSelected = selectedCells.has(cellId);
                              const isSticky = ci < 2;
                              const stickyLeft = ci === 0 ? "left-10" : "left-[calc(2.5rem+110px)]";

                              let displayVal: string;
                              if ((PCT_COLS.has(col) || IMS_PCT_COLS.has(col)) && typeof val === "number") {
                                displayVal = `${(val * 100).toFixed(1)}%`;
                              } else if (typeof val === "number") {
                                displayVal = val.toLocaleString(undefined, { maximumFractionDigits: 2 });
                              } else {
                                displayVal = String(val ?? "");
                              }

                              const highlight = getCellHighlight(col, val);

                              const handleClick = (e: React.MouseEvent) => {
                                e.stopPropagation();
                                if (e.shiftKey && anchorCell) {
                                  // Range select: gather all cells from anchor to current
                                  const anchorRowIdx = displayRows.findIndex(r => String(r.row_num) === anchorCell.rowIdx.toString());
                                  const rMin = Math.min(anchorRowIdx, rowIdx);
                                  const rMax = Math.max(anchorRowIdx, rowIdx);
                                  const cMin = Math.min(visibleCols.indexOf(anchorCell.colLetter), ci);
                                  const cMax = Math.max(visibleCols.indexOf(anchorCell.colLetter), ci);
                                  const newSet = new Set(selectedCells);
                                  for (let ri = rMin; ri <= rMax; ri++) {
                                    for (let ci2 = cMin; ci2 <= cMax; ci2++) {
                                      newSet.add(`${displayRows[ri].row_num}:${visibleCols[ci2]}`);
                                    }
                                  }
                                  setSelectedCells(newSet);
                                } else if (e.ctrlKey || e.metaKey) {
                                  // Toggle individual cell
                                  const newSet = new Set(selectedCells);
                                  if (newSet.has(cellId)) newSet.delete(cellId); else newSet.add(cellId);
                                  setSelectedCells(newSet);
                                  setAnchorCell({ rowIdx: row.row_num, colLetter: col });
                                } else {
                                  // Single click: highlight entire row
                                  const newSet = new Set<string>();
                                  visibleCols.forEach(c => newSet.add(`${rowKey}:${c}`));
                                  setSelectedCells(newSet);
                                  setAnchorCell({ rowIdx: row.row_num, colLetter: col });
                                  handleCellClick(rowIdx, col, cell);
                                }
                              };

                              return (
                                <td key={col}
                                  onClick={handleClick}
                                  className={`px-3 py-1.5 border-r border-gray-800/40 cursor-pointer select-none transition whitespace-nowrap
                    ${isSticky ? `sticky ${stickyLeft} z-10` : ""}
                    ${isFormula ? "text-emerald-400 font-mono" : "text-gray-200"}
                    ${isCellSelected ? "!bg-blue-600/30 ring-1 ring-inset ring-blue-500/60" : isFormula ? "bg-gray-900/20" : isSticky ? "bg-[#0c1424]/80" : ""}
                    ${highlight}
                    ${isEditing(col) ? "ring-2 ring-blue-500 ring-inset bg-blue-950/30" : ""}`}>
                                  {displayVal}
                                </td>
                              );
                            })}
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
                <p className="text-[10px] text-gray-600 mt-2 text-right">{displayRows.length} nodes hiển thị</p>
              </div>
            </div>
          )}

          {/* ══ TAB: DASHBOARD ════════════════════════════════════════════════ */}
          {activeTab === "dashboard" && (
            <div className="space-y-6">
              {/* Node filter with search */}
              <div className="glass-card rounded-2xl p-4 flex items-start gap-4 flex-wrap" style={{ position: "relative", zIndex: 50 }}>
                <div className="flex items-center gap-2 pt-1 shrink-0">
                  <Filter className="w-4 h-4 text-blue-400" />
                  <span className="text-sm font-semibold text-gray-300">Lọc Node:</span>
                </div>
                <div className="relative" style={{ zIndex: 51 }}>
                  <button onClick={() => { setNodeFilterOpen(v => !v); setNodeSearch(""); }}
                    className="px-4 py-1.5 bg-gray-800 hover:bg-gray-700 text-gray-300 rounded-xl text-sm flex items-center gap-2 border border-gray-700 transition">
                    {selectedNodes.length === 0 ? "Tất cả Nodes" : `${selectedNodes.length} / ${allNodeNames.length} node`}
                    <ChevronDown className={`w-4 h-4 transition-transform ${nodeFilterOpen ? "rotate-180" : ""}`} />
                  </button>
                  {nodeFilterOpen && (
                    <div className="absolute top-full mt-2 left-0 bg-[#0c1424] border border-gray-700/80 rounded-2xl shadow-2xl w-72" style={{ zIndex: 9999 }}>
                      {/* Search input */}
                      <div className="p-3 border-b border-gray-800">
                        <div className="relative">
                          <Filter className="absolute left-2.5 top-1/2 -translate-y-1/2 w-3.5 h-3.5 text-gray-500" />
                          <input
                            autoFocus
                            value={nodeSearch}
                            onChange={e => setNodeSearch(e.target.value)}
                            placeholder="Tìm kiếm node..."
                            className="w-full bg-black/40 border border-gray-800 rounded-xl pl-8 pr-3 py-1.5 text-xs text-gray-300 focus:outline-none focus:border-blue-500"
                          />
                        </div>
                      </div>
                      {/* Actions */}
                      <div className="flex justify-between px-3 py-2 border-b border-gray-800">
                        <button onClick={() => {
                          const filtered = allNodeNames.filter(n => n.toLowerCase().includes(nodeSearch.toLowerCase()));
                          setSelectedNodes(prev => Array.from(new Set([...prev, ...filtered])));
                        }} className="text-[10px] text-blue-400 hover:text-blue-300 font-semibold">Chọn tất cả</button>
                        <button onClick={() => setSelectedNodes([])} className="text-[10px] text-gray-400 hover:text-gray-200">Bỏ chọn tất cả</button>
                        <button onClick={() => setNodeFilterOpen(false)} className="text-[10px] text-gray-500 hover:text-gray-300">Đóng</button>
                      </div>
                      {/* List */}
                      <div className="max-h-56 overflow-y-auto p-2">
                        {allNodeNames
                          .filter(n => n.toLowerCase().includes(nodeSearch.toLowerCase()))
                          .map(n => (
                            <label key={n} className="flex items-center gap-2.5 px-2.5 py-1.5 rounded-xl hover:bg-gray-800/60 cursor-pointer">
                              <input type="checkbox" checked={selectedNodes.includes(n)}
                                onChange={e => setSelectedNodes(v => e.target.checked ? [...v, n] : v.filter(x => x !== n))}
                                className="accent-blue-500 w-3.5 h-3.5" />
                              <span className="text-sm text-gray-300">{n}</span>
                            </label>
                          ))
                        }
                        {allNodeNames.filter(n => n.toLowerCase().includes(nodeSearch.toLowerCase())).length === 0 && (
                          <p className="text-xs text-gray-600 text-center py-4">Không tìm thấy node nào</p>
                        )}
                      </div>
                    </div>
                  )}
                </div>
                {/* Selected tags */}
                {selectedNodes.length > 0 && (
                  <div className="flex flex-wrap gap-1.5 flex-1">
                    {selectedNodes.map(n => (
                      <span key={n} className="px-2 py-0.5 bg-blue-600/20 text-blue-300 border border-blue-500/20 rounded-full text-xs flex items-center gap-1">
                        {n}
                        <button onClick={() => setSelectedNodes(v => v.filter(x => x !== n))} className="hover:text-red-400 transition ml-0.5"><X className="w-2.5 h-2.5" /></button>
                      </span>
                    ))}
                  </div>
                )}
              </div>

              {/* ── Area filter checkboxes ── */}
              {allAreas.length > 0 && (
                <div className="glass-card rounded-2xl px-4 py-3 flex items-center gap-4 flex-wrap">
                  <div className="flex items-center gap-2 shrink-0">
                    <SlidersHorizontal className="w-3.5 h-3.5 text-violet-400" />
                    <span className="text-sm font-semibold text-gray-300">Lọc theo Area:</span>
                  </div>
                  <div className="flex flex-wrap gap-2">
                    {allAreas.map(area => {
                      const isActive = selectedAreas.includes(area);
                      return (
                        <label
                          key={area}
                          className={`flex items-center gap-2 px-3 py-1.5 rounded-xl border cursor-pointer text-xs font-semibold transition-all select-none
                            ${isActive
                              ? "bg-violet-600/25 border-violet-500/60 text-violet-200"
                              : "bg-gray-800/60 border-gray-700/60 text-gray-400 hover:border-violet-500/40 hover:text-gray-200"}`}
                        >
                          <input
                            type="checkbox"
                            checked={isActive}
                            onChange={e =>
                              setSelectedAreas(prev =>
                                e.target.checked ? [...prev, area] : prev.filter(a => a !== area)
                              )
                            }
                            className="accent-violet-500 w-3.5 h-3.5"
                          />
                          {area}
                          <span className={`text-[10px] font-normal ${isActive ? "text-violet-400" : "text-gray-600"}`}>
                            ({rows.filter(r => String((r["D"] as CellInfo)?.value) === area).length})
                          </span>
                        </label>
                      );
                    })}
                    {selectedAreas.length > 0 && (
                      <button
                        onClick={() => setSelectedAreas([])}
                        className="px-2.5 py-1.5 rounded-xl border border-gray-700/60 bg-gray-800/40 text-gray-500 hover:text-gray-300 text-xs transition"
                      >
                        Bỏ lọc area
                      </button>
                    )}
                  </div>
                </div>
              )}

              {/* ── HW/NFVI Site filter checkboxes ── */}
              {allHwSites.length > 0 && (
                <div className="glass-card rounded-2xl px-4 py-3 flex items-center gap-4 flex-wrap">
                  <div className="flex items-center gap-2 shrink-0">
                    <Filter className="w-3.5 h-3.5 text-cyan-400" />
                    <span className="text-sm font-semibold text-gray-300">Lọc theo HW/NFVI Site:</span>
                  </div>
                  <div className="flex flex-wrap gap-2">
                    {allHwSites.map(site => {
                      const isActive = selectedHwSites.includes(site);
                      return (
                        <label
                          key={site}
                          className={`flex items-center gap-2 px-3 py-1.5 rounded-xl border cursor-pointer text-xs font-semibold transition-all select-none
                            ${isActive
                              ? "bg-cyan-600/20 border-cyan-500/60 text-cyan-200"
                              : "bg-gray-800/60 border-gray-700/60 text-gray-400 hover:border-cyan-500/40 hover:text-gray-200"}`}
                        >
                          <input
                            type="checkbox"
                            checked={isActive}
                            onChange={e =>
                              setSelectedHwSites(prev =>
                                e.target.checked ? [...prev, site] : prev.filter(s => s !== site)
                              )
                            }
                            className="accent-cyan-500 w-3.5 h-3.5"
                          />
                          {site}
                          <span className={`text-[10px] font-normal ${isActive ? "text-cyan-400" : "text-gray-600"}`}>
                            ({rows.filter(r => String((r["AM"] as CellInfo)?.value) === site).length})
                          </span>
                        </label>
                      );
                    })}
                    {selectedHwSites.length > 0 && (
                      <button
                        onClick={() => setSelectedHwSites([])}
                        className="px-2.5 py-1.5 rounded-xl border border-gray-700/60 bg-gray-800/40 text-gray-500 hover:text-gray-300 text-xs transition"
                      >
                        Bỏ lọc site
                      </button>
                    )}
                  </div>
                </div>
              )}

              {/* ── 3 Merged Charts: Before + After side-by-side ── */}
              <div className="space-y-6">
                {/* Chart 1: Bear */}
                <ChartCard
                  title="Bear Total sử dụng · Bear sử dụng DK / License Bear UCTT / Total IPv4 (v-internet)"
                  data={bearChartData}
                  bars={[
                    { key: "Bear Total sử dụng", color: COLORS.bear, type: "before" },
                    { key: "Bear sử dụng DK", color: COLORS.bearDK, type: "after" },
                    { key: "License Bear UCTT", color: COLORS.licBear, type: "ref" },
                    { key: "Total IPv4 (v-internet)", color: COLORS.ipv4, type: "ref" },
                  ]}
                />

                {/* Chart 2: IMS */}
                <ChartCard
                  title="Bear IMS · Bearer IMS dự kiến / Total IPv4 (IMS)"
                  data={imsChartData}
                  bars={[
                    { key: "Bear IMS", color: COLORS.bearIMS, type: "before" },
                    { key: "Bearer IMS dự kiến", color: COLORS.imsdk, type: "after" },
                    { key: "Total IPv4 (IMS)", color: COLORS.ipv4IMS, type: "ref" },
                  ]}
                />

                {/* Chart 3: Throughput */}
                <ChartCard
                  title="Throughput · Throughput dự kiến / License Throughput VHKT"
                  data={thruChartData}
                  bars={[
                    { key: "Throughput", color: COLORS.thru, type: "before" },
                    { key: "Throughput dự kiến", color: COLORS.thruDK, type: "after" },
                    { key: "License Throughput VHKT", color: COLORS.licThru, type: "ref" },
                  ]}
                />

                {/* Chart 4: % Tải IMS / License */}
                <div className="glass-card rounded-3xl p-5">
                  <div className="flex items-start justify-between mb-3 gap-4">
                    <h4 className="text-sm font-bold text-gray-300">% Tải IMS / License: % Tải IMS trước và % Tải IMS sau</h4>
                    <div className="flex items-center gap-3 shrink-0 text-[10px] text-gray-500">
                      <span className="flex items-center gap-1">
                        <span className="inline-block w-3 h-2 rounded-sm" style={{ background: "#f97316" }} />
                        % Tải IMS trước
                      </span>
                      <span className="flex items-center gap-1">
                        <span className="inline-block w-3 h-2 rounded-sm opacity-75" style={{ background: "#a78bfa" }} />
                        % Tải IMS sau
                      </span>
                      <span className="flex items-center gap-1">
                        <span className="inline-block w-3 h-2 rounded-sm" style={{ background: "#ef4444" }} />
                        Ngưỡng ({overloadThreshold}%)
                      </span>
                    </div>
                  </div>
                  <div className="h-[420px]">
                    <ResponsiveContainer width="100%" height="100%">
                      <BarChart data={imsPctChartData} margin={{ top: 5, right: 20, left: 10, bottom: 70 }}>
                        <CartesianGrid strokeDasharray="3 3" stroke="#1e293b" />
                        <XAxis dataKey="name" stroke="#6b7280" tick={{ fontSize: 11 }} angle={-35} textAnchor="end" interval={0} />
                        <YAxis stroke="#6b7280" tick={{ fontSize: 10 }} tickFormatter={v => `${(v * 100).toFixed(0)}%`} domain={[0, 'auto']} />
                        <Tooltip
                          contentStyle={tooltipStyle}
                          formatter={(value: any, name: any) => [
                            `${(Number(value) * 100).toFixed(1)}%`,
                            name
                          ]}
                        />
                        <Legend wrapperStyle={{ paddingTop: "14px", fontSize: "11px" }} />
                        <ReferenceLine y={overloadThreshold / 100} stroke="#ef4444" strokeDasharray="5 3" strokeWidth={1.5} label={{ value: `${overloadThreshold}%`, fill: "#ef4444", fontSize: 10, position: "insideTopRight" }} />
                        <Bar dataKey="% Tải IMS trước" fill="#f97316" radius={[3, 3, 0, 0]} />
                        <Bar dataKey="% Tải IMS sau" fill="#a78bfa" radius={[3, 3, 0, 0]} fillOpacity={0.8} />
                      </BarChart>
                    </ResponsiveContainer>
                  </div>
                </div>

                {/* Alert summary table */}
                <div className="glass-card rounded-2xl p-5">
                  <p className="text-xs font-bold text-gray-400 uppercase tracking-wider mb-3 flex items-center gap-2">
                    <AlertTriangle className="w-3.5 h-3.5 text-amber-400" /> Cảnh báo vượt ngưỡng ({overloadThreshold}%)
                  </p>
                  <div className="grid grid-cols-1 lg:grid-cols-3 gap-4">
                    {/* Bear */}
                    <div className="space-y-2">
                      <p className="text-[10px] font-bold text-blue-400 uppercase tracking-wider mb-1">Bear</p>
                      {(() => {
                        const rows = filteredDashRows.filter(r => {
                          const g = (r["G"] as CellInfo)?.value as number || 1;
                          const before = (r["L"] as CellInfo)?.value as number || 0;
                          const after = (r["T"] as CellInfo)?.value as number || 0;
                          return g > 0 && ((before / g) * 100 > overloadThreshold || (after / g) * 100 > overloadThreshold);
                        });
                        if (rows.length === 0) return <p className="text-[10px] text-emerald-400 flex items-center gap-1"><CheckCircle2 className="w-3 h-3" /> Bình thường</p>;
                        return rows.map((r, i) => {
                          const node = String((r["B"] as CellInfo)?.value || "");
                          const g = (r["G"] as CellInfo)?.value as number || 1;
                          const bef = g > 0 ? ((r["L"] as CellInfo)?.value as number || 0) / g * 100 : 0;
                          const aft = g > 0 ? ((r["T"] as CellInfo)?.value as number || 0) / g * 100 : 0;
                          const delta = aft - bef;
                          return (
                            <div key={i} className="flex justify-between items-center text-[10px] bg-gray-900/40 rounded-lg px-2.5 py-1.5">
                              <span className="font-semibold text-gray-300">{node}</span>
                              <div className="flex items-center gap-2">
                                <span className={bef > overloadThreshold ? "text-red-400 font-bold" : "text-gray-500"}>{bef.toFixed(1)}%</span>
                                <span className="text-gray-600">→</span>
                                <span className={aft > overloadThreshold ? "text-orange-400 font-bold" : "text-emerald-400"}>{aft.toFixed(1)}%</span>
                                <span className={`font-bold ${delta > 0 ? "text-red-400" : "text-emerald-400"}`}>
                                  {delta > 0 ? "▲" : "▼"}{Math.abs(delta).toFixed(1)}%
                                </span>
                              </div>
                            </div>
                          );
                        });
                      })()}
                    </div>
                    {/* IMS */}
                    <div className="space-y-2">
                      <p className="text-[10px] font-bold text-pink-400 uppercase tracking-wider mb-1">Bear IMS</p>
                      {(() => {
                        const rows = filteredDashRows.filter(r => {
                          const cap = (r["I2"] as CellInfo)?.value as number || 1;
                          const bef = (r["M"] as CellInfo)?.value as number || 0;
                          const aft = (r["V"] as CellInfo)?.value as number || 0;
                          return cap > 0 && ((bef / cap) * 100 > overloadThreshold || (aft / cap) * 100 > overloadThreshold);
                        });
                        if (rows.length === 0) return <p className="text-[10px] text-emerald-400 flex items-center gap-1"><CheckCircle2 className="w-3 h-3" /> Bình thường</p>;
                        return rows.map((r, i) => {
                          const node = String((r["B"] as CellInfo)?.value || "");
                          const cap = (r["I2"] as CellInfo)?.value as number || 1;
                          const bef = cap > 0 ? ((r["M"] as CellInfo)?.value as number || 0) / cap * 100 : 0;
                          const aft = cap > 0 ? ((r["V"] as CellInfo)?.value as number || 0) / cap * 100 : 0;
                          const delta = aft - bef;
                          return (
                            <div key={i} className="flex justify-between items-center text-[10px] bg-gray-900/40 rounded-lg px-2.5 py-1.5">
                              <span className="font-semibold text-gray-300">{node}</span>
                              <div className="flex items-center gap-2">
                                <span className={bef > overloadThreshold ? "text-red-400 font-bold" : "text-gray-500"}>{bef.toFixed(1)}%</span>
                                <span className="text-gray-600">→</span>
                                <span className={aft > overloadThreshold ? "text-orange-400 font-bold" : "text-emerald-400"}>{aft.toFixed(1)}%</span>
                                <span className={`font-bold ${delta > 0 ? "text-red-400" : "text-emerald-400"}`}>
                                  {delta > 0 ? "▲" : "▼"}{Math.abs(delta).toFixed(1)}%
                                </span>
                              </div>
                            </div>
                          );
                        });
                      })()}
                    </div>
                    {/* Throughput */}
                    <div className="space-y-2">
                      <p className="text-[10px] font-bold text-cyan-400 uppercase tracking-wider mb-1">Throughput</p>
                      {(() => {
                        const rows = filteredDashRows.filter(r => {
                          const h = (r["H"] as CellInfo)?.value as number || 1;
                          const bef = (r["N"] as CellInfo)?.value as number || 0;
                          const aft = (r["U"] as CellInfo)?.value as number || 0;
                          return h > 0 && ((bef / h) * 100 > overloadThreshold || (aft / h) * 100 > overloadThreshold);
                        });
                        if (rows.length === 0) return <p className="text-[10px] text-emerald-400 flex items-center gap-1"><CheckCircle2 className="w-3 h-3" /> Bình thường</p>;
                        return rows.map((r, i) => {
                          const node = String((r["B"] as CellInfo)?.value || "");
                          const h = (r["H"] as CellInfo)?.value as number || 1;
                          const bef = h > 0 ? ((r["N"] as CellInfo)?.value as number || 0) / h * 100 : 0;
                          const aft = h > 0 ? ((r["U"] as CellInfo)?.value as number || 0) / h * 100 : 0;
                          const delta = aft - bef;
                          return (
                            <div key={i} className="flex justify-between items-center text-[10px] bg-gray-900/40 rounded-lg px-2.5 py-1.5">
                              <span className="font-semibold text-gray-300">{node}</span>
                              <div className="flex items-center gap-2">
                                <span className={bef > overloadThreshold ? "text-red-400 font-bold" : "text-gray-500"}>{bef.toFixed(1)}%</span>
                                <span className="text-gray-600">→</span>
                                <span className={aft > overloadThreshold ? "text-orange-400 font-bold" : "text-emerald-400"}>{aft.toFixed(1)}%</span>
                                <span className={`font-bold ${delta > 0 ? "text-red-400" : "text-emerald-400"}`}>
                                  {delta > 0 ? "▲" : "▼"}{Math.abs(delta).toFixed(1)}%
                                </span>
                              </div>
                            </div>
                          );
                        });
                      })()}
                    </div>
                  </div>
                </div>
              </div>
            </div>
          )}
        </div>
      </main>
    </div>
  );
}
