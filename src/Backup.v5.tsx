/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useEffect, useMemo } from 'react';
import { motion, AnimatePresence } from 'motion/react';
import { 
  Search, 
  Plus, 
  Trash2, 
  FileText, 
  ChevronRight, 
  ChevronLeft, 
  CheckCircle2, 
  AlertCircle,
  Loader2,
  Download,
  Package,
  Settings2,
  RefreshCw,
  ChevronDown
} from 'lucide-react';
import Papa from 'papaparse';
import { jsPDF } from 'jspdf';
import autoTable from 'jspdf-autotable';
import { THSarabunNew_normal, THSarabunNew_bold } from "../fonts";
import { 
  Document, 
  Packer, 
  Paragraph, 
  Table, 
  TableRow, 
  TableCell, 
  WidthType, 
  AlignmentType, 
  HeadingLevel,
  TextRun,
  VerticalAlign,
  BorderStyle,
  ImageRun,
  Footer,
  Header,
  PageNumber,
} from 'docx';
import { saveAs } from 'file-saver';
import { cn } from '@/src/lib/utils';
import { Material, AssessmentItem, GOOGLE_SHEET_ID } from './types';

// Add type for jsPDF with autotable
interface jsPDFWithAutoTable extends jsPDF {
  autoTable: (options: any) => jsPDF;
}

export default function App() {
  const [materials, setMaterials] = useState<Material[]>([]);
  const [loading, setLoading] = useState(true);
  const [exporting, setExporting] = useState(false);
  const [error, setError] = useState<string | null>(null);
  
  // Selection State
  const [searchQuery, setSearchQuery] = useState('');
  const [selectedMaterial, setSelectedMaterial] = useState<Material | null>(null);
  const [quantity, setQuantity] = useState<number>(1);
  const [quantityInput, setQuantityInput] = useState<string>("1");

  useEffect(() => {
    const currentInputNumeric = parseInt(quantityInput.replace(/[^0-9]/g, ''), 10) || 0;
    if (currentInputNumeric !== quantity) {
      setQuantityInput(quantity.toLocaleString());
    }
  }, [quantity]);
  
  // Assessment List
  const [items, setItems] = useState<AssessmentItem[]>([]);
  const [showClearConfirm, setShowClearConfirm] = useState(false);

  // Selected Sheet State
  const sheetOptions = [
    'ราคาพัสดุแบบสำเร็จรูป ครั้งที่ 3/2569',
    'ราคาพัสดุแบบสำเร็จรูป ครั้งที่ 2/2569',
    'ราคาพัสดุแบบสำเร็จรูป ครั้งที่ 1/2569',
    'ราคาพัสดุแบบสำเร็จรูป ครั้งที่ 3/2568',
    'ราคาพัสดุแบบสำเร็จรูป ครั้งที่ 2/2568',
    'ราคาพัสดุแบบสำเร็จรูป ครั้งที่ 1/2568',
  ];

  const SHEET_GIDS: Record<string, string> = {
    'ราคาพัสดุแบบสำเร็จรูป ครั้งที่ 3/2569': '0',
    'ราคาพัสดุแบบสำเร็จรูป ครั้งที่ 2/2569': '105388871',
    'ราคาพัสดุแบบสำเร็จรูป ครั้งที่ 1/2569': '1972078162',
    'ราคาพัสดุแบบสำเร็จรูป ครั้งที่ 3/2568': '1688423232',
    'ราคาพัสดุแบบสำเร็จรูป ครั้งที่ 2/2568': '1052539521',
    'ราคาพัสดุแบบสำเร็จรูป ครั้งที่ 1/2568': '1417798826',
  };

  const [selectedSheet, setSelectedSheet] = useState<string>('ราคาพัสดุแบบสำเร็จรูป ครั้งที่ 3/2569');

  // Fetch data from Google Sheet
  const fetchData = async (sheetName?: string) => {
    const targetSheet = sheetName !== undefined ? sheetName : selectedSheet;
    const gid = SHEET_GIDS[targetSheet];
    try {
      setLoading(true);
      setError(null);

      const url = gid !== undefined
        ? `https://docs.google.com/spreadsheets/d/${GOOGLE_SHEET_ID}/export?format=csv&gid=${gid}&_t=${Date.now()}`
        : `https://docs.google.com/spreadsheets/d/${GOOGLE_SHEET_ID}/export?format=csv&sheet=${encodeURIComponent(targetSheet)}&_t=${Date.now()}`;
      const response = await fetch(url, { cache: 'no-store' });
      
      if (!response.ok) {
        throw new Error('Network response was not ok');
      }
      
      const csvText = await response.text();
      
      Papa.parse(csvText, {
        header: false,
        skipEmptyLines: true,
        complete: (results) => {
          if (results.data.length <= 1) {
            setError('ไม่พบข้อมูลใน Google Sheet หรือไฟล์ว่างเปล่า');
            setLoading(false);
            return;
          }

          const dataRows = results.data.slice(1);

          const mappedData: Material[] = dataRows
            .filter((row: any) => row[2])
            .map((row: any) => {
              const id = String(row[1] || '').trim();
              const name = String(row[2] || '').trim();
              const unit = String(row[3] || 'หน่วย').trim();
              
              const parsePrice = (val: any) => {
                if (!val) return 0;
                return parseFloat(String(val).replace(/,/g, '').replace(/฿/g, '').trim()) || 0;
              };

              const priceDamaged = parsePrice(row[4]);
              const priceReusable = parsePrice(row[5]);
              
              return { id, name, unit, priceDamaged, priceReusable };
            });
          
          if (mappedData.length === 0) {
            setError('ไม่สามารถดึงข้อมูลพัสดุได้ กรุณาตรวจสอบรูปแบบข้อมูลใน Sheet');
          } else {
            setMaterials(mappedData);

            // Update currently selected material in Step 1 if active
            setSelectedMaterial((prevMat) => {
              if (!prevMat) return null;
              const updated = mappedData.find((m) => m.id === prevMat.id);
              return updated || prevMat;
            });

            // Update items in current assessment list with new prices if sheet changes
            setItems((prevItems) => {
              if (prevItems.length === 0) return prevItems;
              return prevItems.map((item) => {
                const updatedMaterial = mappedData.find((m) => m.id === item.material.id) || item.material;
                const unitPrice = item.status === 'damaged' ? updatedMaterial.priceDamaged : updatedMaterial.priceReusable;
                return {
                  ...item,
                  material: updatedMaterial,
                  totalPrice: item.quantity * unitPrice,
                };
              });
            });
          }
          setLoading(false);
        },
        error: (err: any) => {
          console.error('Parsing error:', err);
          setError('ไม่สามารถประมวลผลข้อมูล CSV ได้');
          setLoading(false);
        }
      });
    } catch (err) {
      console.error('Fetch error:', err);
      setError('ไม่สามารถเชื่อมต่อกับ Google Sheet ได้ กรุณาตรวจสอบการแชร์ไฟล์');
      setLoading(false);
    }
  };

  useEffect(() => {
    fetchData(selectedSheet);
  }, [selectedSheet]);

  const filteredMaterials = useMemo(() => {
    if (!searchQuery) return [];
    return materials.filter(m => 
      m.name.toLowerCase().includes(searchQuery.toLowerCase()) || 
      m.id.toLowerCase().includes(searchQuery.toLowerCase())
    ).slice(0, 10);
  }, [materials, searchQuery]);

  const handleStatusSelect = (selectedStatus: 'damaged' | 'reusable') => {
    if (!selectedMaterial) return;
    
    const currentPrice = selectedStatus === 'damaged' ? selectedMaterial.priceDamaged : selectedMaterial.priceReusable;
    
    const newItem: AssessmentItem = {
      material: selectedMaterial,
      quantity,
      status: selectedStatus,
      totalPrice: currentPrice * quantity
    };
    
    setItems([...items, newItem]);
    resetForm();
  };

  const resetForm = () => {
    setSelectedMaterial(null);
    setSearchQuery('');
    setQuantity(1);
  };

  const removeItem = (index: number) => {
    setItems(items.filter((_, i) => i !== index));
  };

  // Helper to group items by material name
  const groupItems = (itemsList: AssessmentItem[]) => {
    const grouped: { [key: string]: AssessmentItem } = {};
    itemsList.forEach(item => {
      const key = item.material.name;
      if (grouped[key]) {
        grouped[key] = {
          ...grouped[key],
          quantity: grouped[key].quantity + item.quantity,
          totalPrice: grouped[key].totalPrice + item.totalPrice
        };
      } else {
        grouped[key] = { ...item };
      }
    });
    return Object.values(grouped);
  };

  const totalAmount = items.reduce((sum, item) => sum + item.totalPrice, 0);

  const exportWord = async () => {
    // Helper to fetch image and convert to ArrayBuffer
    const fetchImage = async (url: string) => {
      try {
        const response = await fetch(url);
        const blob = await response.blob();
        return await blob.arrayBuffer();
      } catch (error) {
        console.error("Error fetching logo:", error);
        return null;
      }
    };

    const logoBuffer = await fetchImage("https://img1.pic.in.th/images/PEA-02-Thai-Logo.md.jpg");

    const damagedItems = groupItems(items.filter(i => i.status === 'damaged'));
    const reusableItems = groupItems(items.filter(i => i.status === 'reusable'));
    
    const totalItems = damagedItems.length + reusableItems.length;
    
    const damagedCount = damagedItems.length;
    const damagedTotal = damagedItems.reduce((sum, i) => sum + i.totalPrice, 0);
    
    const reusableCount = reusableItems.length;
    const reusableTotal = reusableItems.reduce((sum, i) => sum + i.totalPrice, 0);

    const createTableHeader = () => new TableRow({
      children: [
        new TableCell({ width: { size: 10, type: WidthType.PERCENTAGE }, children: [new Paragraph({ text: "รายการ", alignment: AlignmentType.CENTER, style: "bold" })] }),
        new TableCell({ width: { size: 50, type: WidthType.PERCENTAGE }, children: [new Paragraph({ text: "ชื่อพัสดุ", alignment: AlignmentType.CENTER, style: "bold" })] }),
        new TableCell({ width: { size: 10, type: WidthType.PERCENTAGE }, children: [new Paragraph({ text: "จำนวน", alignment: AlignmentType.CENTER, style: "bold" })] }),
        new TableCell({ width: { size: 10, type: WidthType.PERCENTAGE }, children: [new Paragraph({ text: "หน่วย", alignment: AlignmentType.CENTER, style: "bold" })] }),
        new TableCell({ width: { size: 20, type: WidthType.PERCENTAGE }, children: [new Paragraph({ text: "ราคา", alignment: AlignmentType.CENTER, style: "bold" })] }),
      ],
    });

    const damagedTableRows = damagedItems.map((item, index) => (
      new TableRow({
        children: [
          new TableCell({ children: [new Paragraph({ text: (index + 1).toString(), alignment: AlignmentType.CENTER })] }),
          new TableCell({ children: [new Paragraph({ text: item.material.name })] }),
          new TableCell({ children: [new Paragraph({ text: item.quantity.toLocaleString(), alignment: AlignmentType.CENTER })] }),
          new TableCell({ children: [new Paragraph({ text: item.material.unit, alignment: AlignmentType.CENTER })] }),
          new TableCell({ children: [new Paragraph({ text: item.totalPrice.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 }), alignment: AlignmentType.RIGHT })] }),
        ],
      })
    ));

    const reusableTableRows = reusableItems.map((item, index) => (
      new TableRow({
        children: [
          new TableCell({ children: [new Paragraph({ text: (index + 1).toString(), alignment: AlignmentType.CENTER })] }),
          new TableCell({ children: [new Paragraph({ text: item.material.name })] }),
          new TableCell({ children: [new Paragraph({ text: item.quantity.toLocaleString(), alignment: AlignmentType.CENTER })] }),
          new TableCell({ children: [new Paragraph({ text: item.material.unit, alignment: AlignmentType.CENTER })] }),
          new TableCell({ children: [new Paragraph({ text: item.totalPrice.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 }), alignment: AlignmentType.RIGHT })] }),
        ],
      })
    ));

    const createEmptyRows = (count: number) => {
      const rows = [];
      for (let i = 0; i < count; i++) {
        rows.push(new TableRow({
          children: [
            new TableCell({ children: [new Paragraph({ text: "" })] }),
            new TableCell({ children: [new Paragraph({ text: "" })] }),
            new TableCell({ children: [new Paragraph({ text: "" })] }),
            new TableCell({ children: [new Paragraph({ text: "" })] }),
            new TableCell({ children: [new Paragraph({ text: "" })] }),
          ],
        }));
      }
      return rows;
    };

    const damagedTableRowsWithEmpty = [...damagedTableRows, ...createEmptyRows(Math.max(0, 5 - damagedItems.length))];
    const reusableTableRowsWithEmpty = [...reusableTableRows, ...createEmptyRows(Math.max(0, 5 - reusableItems.length))];

    const doc = new Document({
      styles: {
        default: {
          document: {
            run: {
              font: "TH SarabunPSK",
              size: 32, // 16pt
            },
          },
        },
      },
      sections: [{
        properties: {
          titlePage: true,
          page: {
            margin: {
              top: 1417, // 2.5 cm
              bottom: 1417,
              left: 1701, // 3 cm
              right: 1134, // 2 cm
            },
          },
        },
        headers: {
          first: new Header({ children: [] }),
          default: new Header({
            children: [
              new Paragraph({
                children: [
                  new TextRun({ text: "- " }),
                  new TextRun({ children: [PageNumber.CURRENT] }),
                  new TextRun({ text: " -" })
                ],
                alignment: AlignmentType.CENTER,
              }),
            ],
          }),
        },
        footers: {
          first: new Footer({ children: [] }),
          default: new Footer({
            children: [
              new Paragraph({
                text: "หน่วยงาน",
              }),
              new Paragraph({
                text: "โทร. ...........................................................",
              }),
            ],
          }),
        },
        children: [
          // Logo
          ...(logoBuffer ? [
            new Paragraph({
              children: [
                new ImageRun({
                  data: logoBuffer,
                  type: "jpg",
                  transformation: {
                    width: 157,
                    height: 132,
                  },
                }),
              ],
              alignment: AlignmentType.LEFT,
              spacing: { after: 100 },
            })
          ] : []),
          
          // Header Table for From/To and No/Date
          new Table({
            width: { size: 9072, type: WidthType.DXA }, // 16 cm total width
            borders: {
              top: { style: BorderStyle.NONE },
              bottom: { style: BorderStyle.NONE },
              left: { style: BorderStyle.NONE },
              right: { style: BorderStyle.NONE },
              insideHorizontal: { style: BorderStyle.NONE },
              insideVertical: { style: BorderStyle.NONE },
            },
            rows: [
              new TableRow({
                children: [
                  new TableCell({ 
                    width: { size: 4536, type: WidthType.DXA }, // 8 cm (3cm margin + 8cm = 11cm from edge)
                    children: [new Paragraph({ text: "จาก                                                                      " })] 
                  }),
                  new TableCell({ 
                    width: { size: 4536, type: WidthType.DXA }, // 8 cm
                    children: [new Paragraph({ text: "ถึง                                                                      " })] 
                  }),
                ],
              }),
              new TableRow({
                children: [
                  new TableCell({ 
                    width: { size: 4536, type: WidthType.DXA }, // 8 cm
                    children: [new Paragraph({ text: "เลขที่                                                                   " })] 
                  }),
                  new TableCell({ 
                    width: { size: 4536, type: WidthType.DXA }, // 8 cm
                    children: [new Paragraph({ text: "วันที่                                                                   " })] 
                  }),
                ],
              }),
            ],
          }),

          new Paragraph({
            children: [new TextRun({ text: "เรื่อง   การประเมินค่าเสียหายที่เกิดขึ้นกับระบบจำหน่าย" })],
            spacing: { before: 200 },
          }),
          new Paragraph({ text: "เรียน                                                                    " }),
          new Paragraph({
            children: [
              new TextRun({ text: "ตามที่ได้ดำเนินการตรวจสอบและประเมินราคาค่าเสียหายเพื่อเรียกร้องจากผู้กระทำละเมิด โดยมีรายละเอียด ดังนี้" }),
            ],
            indent: { firstLine: 1417 }, // 2.5 cm
            spacing: { before: 200 },
            alignment: AlignmentType.THAI_DISTRIBUTE,
          }),
          new Paragraph({ text: "1. เหตุเกิดเมื่อ ...........................................................................................................................", indent: { left: 1417 }, alignment: AlignmentType.THAI_DISTRIBUTE }),
          new Paragraph({ text: "2. สถานที่เกิดเหตุ .....................................................................................................................", indent: { left: 1417 }, alignment: AlignmentType.THAI_DISTRIBUTE }),
          new Paragraph({ text: "3. หมายเลขทะเบียน ................................................................................................................", indent: { left: 1417 }, alignment: AlignmentType.THAI_DISTRIBUTE }),
          new Paragraph({ text: "4. ชื่อผู้ขับขี่ ..............................................................................................................................", indent: { left: 1417 }, alignment: AlignmentType.THAI_DISTRIBUTE }),
          new Paragraph({ text: "   บัตรประชาชนเลขที่ ......................................................................................................", indent: { left: 1417 }, alignment: AlignmentType.THAI_DISTRIBUTE }),
          new Paragraph({ text: "5. ที่อยู่ตามบัตร ........................................................................................................................", indent: { left: 1417 }, alignment: AlignmentType.THAI_DISTRIBUTE }),
          new Paragraph({ text: "   ........................................................................................................... เบอร์โทรศัพท์ ....................................", alignment: AlignmentType.THAI_DISTRIBUTE }),
          new Paragraph({ text: "6. ชื่อ/บริษัท เจ้าของรถยนต์ ..................................... เบอร์โทรศัพท์ ....................................", indent: { left: 1417 }, alignment: AlignmentType.THAI_DISTRIBUTE }),
          new Paragraph({ text: "7. ชื่อ/บริษัท ประกันภัย ............................................ เบอร์โทรศัพท์ ....................................", indent: { left: 1417 }, alignment: AlignmentType.THAI_DISTRIBUTE }),
          new Paragraph({ text: "8. ผู้ลงนามในหนังสือรับสภาพหนี้", indent: { left: 1417 }, alignment: AlignmentType.THAI_DISTRIBUTE }),
          new Paragraph({ text: "   [  ] ผู้ขับขี่    [  ] เจ้าของรถยนต์    [  ] ไม่ยินยอม", indent: { left: 1417 }, alignment: AlignmentType.THAI_DISTRIBUTE }),
          new Paragraph({ text: "9. การแจ้งความร้องทุกข์กับเจ้าหน้าที่ตำรวจ", indent: { left: 1417 }, alignment: AlignmentType.THAI_DISTRIBUTE }),
          new Paragraph({ text: "   [  ] แจ้งเป็นหลักฐาน    [  ] แจ้งความเป็นคดี เนื่องจาก ...........................................", indent: { left: 1417 }, alignment: AlignmentType.THAI_DISTRIBUTE }),
          new Paragraph({ text: "10. กรณีรถยนต์เกี่ยวสายสื่อสารทำให้เกิดความเสียหายกับระบบจำหน่าย", indent: { left: 1417 }, alignment: AlignmentType.THAI_DISTRIBUTE }),
          new Paragraph({ text: "    ชื่อ/บริษัท เจ้าของสายสื่อสาร ............................................ ความสูง ...........................", indent: { left: 1417 }, alignment: AlignmentType.THAI_DISTRIBUTE }),
          new Paragraph({
            children: [
              new TextRun({ text: `11. รายการอุปกรณ์ที่ได้รับความเสียหาย ${totalItems} รายการ คิดเป็นค่าเสียหาย จำนวนเงินทั้งสิ้น ${totalAmount.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })} บาท โดยมีรายละเอียดดังนี้` }),
            ],
            spacing: { before: 400 },
            indent: { firstLine: 1417 }, // 2.5 cm for first line only
            alignment: AlignmentType.THAI_DISTRIBUTE,
            pageBreakBefore: true,
          }),
          new Paragraph({
            text: `11.1 รื้อถอน - ติดตั้งใหม่ ${damagedCount} รายการ เป็นจำนวนเงินทั้งสิ้น ${damagedTotal.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })} บาท`,
            indent: { left: 1701 },
            spacing: { before: 200, after: 100 },
            alignment: AlignmentType.THAI_DISTRIBUTE,
          }),
          new Table({
            width: { size: 100, type: WidthType.PERCENTAGE },
            alignment: AlignmentType.CENTER,
            rows: [createTableHeader(), ...damagedTableRowsWithEmpty],
          }),
          new Paragraph({
            text: `11.2 แผนกซ่อมแซม ${reusableCount} รายการ เป็นจำนวนเงินทั้งสิ้น ${reusableTotal.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })} บาท`,
            indent: { left: 1701 },
            spacing: { before: 400, after: 100 },
            alignment: AlignmentType.THAI_DISTRIBUTE,
          }),
          new Table({
            width: { size: 100, type: WidthType.PERCENTAGE },
            alignment: AlignmentType.CENTER,
            rows: [createTableHeader(), ...reusableTableRowsWithEmpty],
          }),
          new Paragraph({
            text: `จึงเรียนมาเพื่อพิจารณาอนุมัติให้ดำเนินการเบิกอุปกรณ์ไปซ่อมแซมตามรายการดังกล่าว พร้อมทั้งเป็นการเรียกเก็บเงินค่าเสียหายจากผู้กระทำละเมิด เป็นจำนวนเงินทั้งสิ้น ${totalAmount.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })} บาท`,
            indent: { firstLine: 1417 }, // 2.5 cm
            spacing: { before: 600 },
            alignment: AlignmentType.THAI_DISTRIBUTE,
          }),
          new Paragraph({
            text: "(...........................................................)",
            alignment: AlignmentType.CENTER,
            spacing: { before: 1000 },
            indent: { left: 4000 },
          }),
          new Paragraph({
            text: "ตำแหน่ง",
            alignment: AlignmentType.CENTER,
            indent: { left: 4000 },
          }),
        ],
      }],
    });

    const blob = await Packer.toBlob(doc);
    saveAs(blob, "ค่าละเมิด1234_PI InnoTech.docx");
  };

  const exportPDF = async () => {
    setExporting(true);
    try {
      // ---- Setup jsPDF ----
      const pdf = new jsPDF({ orientation: 'portrait', unit: 'mm', format: 'a4' });

      // Register TH SarabunPSK font (using imported Base64 from fonts.ts)
      pdf.addFileToVFS('THSarabunNew.ttf', THSarabunNew_normal);
      pdf.addFont('THSarabunNew.ttf', 'THSarabunNew', 'normal');
      pdf.addFileToVFS('THSarabunNew-Bold.ttf', THSarabunNew_bold);
      pdf.addFont('THSarabunNew-Bold.ttf', 'THSarabunNew', 'bold');

      const setF = (style: 'normal' | 'bold' = 'normal', size = 16) => {
        pdf.setFont('THSarabunNew', style);
        pdf.setFontSize(size);
      };

      const pageW = pdf.internal.pageSize.getWidth();  // 210 mm
      const pageH = pdf.internal.pageSize.getHeight(); // 297 mm
      const mL    = 30;  // left margin 3 cm
      const mR    = 20;  // right margin 2 cm
      const mT    = 25;  // top margin
      const mB    = 25;  // bottom margin
      const cW    = pageW - mL - mR; // content width = 160 mm
      const lh    = 8;   // line height mm

      const fmt = (n: number) =>
        n.toLocaleString('th-TH', { minimumFractionDigits: 2, maximumFractionDigits: 2 });

      // ---- Prepare data ----
      const damagedItems  = groupItems(items.filter(i => i.status === 'damaged'));
      const reusableItems = groupItems(items.filter(i => i.status === 'reusable'));
      const totalItems    = damagedItems.length + reusableItems.length;
      const damagedTotal  = damagedItems.reduce((s, i) => s + i.totalPrice, 0);
      const reusableTotal = reusableItems.reduce((s, i) => s + i.totalPrice, 0);

      // ============================================================
      // PAGE 1
      // ============================================================
      setF('normal', 16);
      let y = mT;

      // Logo (optional — skip gracefully if fetch fails)
      try {
        const logoRes = await fetch('https://img1.pic.in.th/images/PEA-02-Thai-Logo.md.jpg');
        if (logoRes.ok) {
          const logoBlob = await logoRes.blob();
          const logoBase64 = await new Promise<string>((resolve) => {
            const reader = new FileReader();
            reader.onload = () => resolve((reader.result as string).split(',')[1]);
            reader.readAsDataURL(logoBlob);
          });
          pdf.addImage(logoBase64, 'JPEG', mL, y, 41.5, 34.9);
          y += 34.9 + 8;
        }
      } catch {
        // logo is optional — continue without it
      }

      // จาก / ถึง / เลขที่ / วันที่
      setF('normal', 16);
      const halfW = cW / 2;
      pdf.text('จาก                                                                      ', mL, y);
      pdf.text('ถึง                                                                      ', mL + halfW, y);
      y += lh;
      pdf.text('เลขที่                                                                 ', mL, y);
      pdf.text('วันที่                                                                 ', mL + halfW, y);
      y += lh + 2;

      pdf.text('เรื่อง   การประเมินค่าเสียหายที่เกิดขึ้นกับระบบจำหน่าย', mL, y);
      y += lh;
      pdf.text('เรียน                                                                      ', mL, y);
      y += lh + 2;

      // Indent paragraph (2.5 cm = 25 mm)
      const introIndent = 25;
      const introText = 'ตามที่ได้ดำเนินการตรวจสอบและประเมินราคาค่าเสียหายเพื่อเรียกร้องจากผู้กระทำละเมิด โดยมีรายละเอียด ดังนี้';
      
      // Split text to handle first line indent manually or using options
      // jsPDF auto-justify doesn't handle first-line indent well with simple text()
      // We'll split it into the first part and the rest if needed, or use a simpler approach
      const firstLineMax = cW - introIndent;
      const introLines = pdf.splitTextToSize(introText, cW); // Split by full width first to see wrap
      
      // For a more precise look like the image:
      pdf.text('ตามที่ได้ดำเนินการตรวจสอบและประเมินราคาค่าเสียหายเพื่อเรียกร้องจากผู้กระทำละเมิด', mL + introIndent, y);
      y += lh;
      pdf.text('โดยมีรายละเอียด ดังนี้', mL, y);
      y += lh + 2;

      // ข้อ 1-10 (Indented 2.5 cm = 25 mm)
      const itemIndent = 25; // Indent for the numbers (2.5 cm)
      const subItemIndent = 30; // Indent for sub-items like "บัตรประชาชนเลขที่"
      const endX = 180; // Target X coordinate for dots (30mm from right edge of 210mm page)

      const addDotsToTarget = (text: string, x: number, targetX: number) => {
        let currentText = text;
        while (pdf.getTextWidth(currentText) + x < targetX) {
          currentText += '.';
        }
        return currentText;
      };

      // 1-3
      pdf.text(addDotsToTarget('1. เหตุเกิดเมื่อ ', mL + itemIndent, endX), mL + itemIndent, y);
      y += lh;
      pdf.text(addDotsToTarget('2. สถานที่เกิดเหตุ ', mL + itemIndent, endX), mL + itemIndent, y);
      y += lh;
      pdf.text(addDotsToTarget('3. หมายเลขทะเบียน ', mL + itemIndent, endX), mL + itemIndent, y);
      y += lh;
      
      // 4
      pdf.text(addDotsToTarget('4. ชื่อผู้ขับขี่ ', mL + itemIndent, endX), mL + itemIndent, y);
      y += lh;
      pdf.text(addDotsToTarget('บัตรประชาชนเลขที่ ', mL + subItemIndent, endX), mL + subItemIndent, y);
      y += lh;
      
      // 5
      pdf.text(addDotsToTarget('5. ที่อยู่ตามบัตร ', mL + itemIndent, endX), mL + itemIndent, y);
      y += lh;
      
      // Line under 5 with leading dots and phone number
      const lineDotsStartX = mL + 10;
      const phoneX = 125;
      let dotsBeforePhone = '';
      while (pdf.getTextWidth(dotsBeforePhone) + lineDotsStartX < phoneX) {
        dotsBeforePhone += '.';
      }
      pdf.text(dotsBeforePhone, lineDotsStartX, y);
      pdf.text(addDotsToTarget(' เบอร์โทรศัพท์ ', phoneX, endX), phoneX, y);
      y += lh;

      // 6
      pdf.text(addDotsToTarget('6. ชื่อ/บริษัท เจ้าของรถยนต์ ', mL + itemIndent, phoneX), mL + itemIndent, y);
      pdf.text(addDotsToTarget(' เบอร์โทรศัพท์ ', phoneX, endX), phoneX, y);
      y += lh;

      // 7
      pdf.text(addDotsToTarget('7. ชื่อ/บริษัท ประกันภัย ', mL + itemIndent, phoneX), mL + itemIndent, y);
      pdf.text(addDotsToTarget(' เบอร์โทรศัพท์ ', phoneX, endX), phoneX, y);
      y += lh;

      // 8
      pdf.text('8. ผู้ลงนามในหนังสือรับสภาพหนี้', mL + itemIndent, y);
      y += lh;
      pdf.text('[  ] ผู้ขับขี่    [  ] เจ้าของรถยนต์    [  ] ไม่ยินยอม', mL + subItemIndent, y);
      y += lh;

      // 9
      pdf.text('9. การแจ้งความร้องทุกข์กับเจ้าหน้าที่ตำรวจ', mL + itemIndent, y);
      y += lh;
      pdf.text(addDotsToTarget('[  ] แจ้งเป็นหลักฐาน    [  ] แจ้งความเป็นคดี เนื่องจาก ', mL + subItemIndent, endX), mL + subItemIndent, y);
      y += lh;

      // 10
      pdf.text('10. กรณีรถยนต์เกี่ยวสายสื่อสารทำให้เกิดความเสียหายกับระบบจำหน่าย', mL + itemIndent, y);
      y += lh;
      const dotShift = pdf.getTextWidth('.') * 20;
      const shiftedX = phoneX + dotShift;
      pdf.text(addDotsToTarget('ชื่อ/บริษัท เจ้าของสายสื่อสาร ', mL + subItemIndent, shiftedX), mL + subItemIndent, y);
      pdf.text(addDotsToTarget(' ความสูง ', shiftedX, endX), shiftedX, y);
      y += lh;

      // ============================================================
      // PAGE 2
      // ============================================================
      pdf.addPage();
      y = mT;

      // Page number header
      setF('normal', 14);
      pdf.text('- 2 -', pageW / 2, 12, { align: 'center' });
      setF('normal', 16);

      // ข้อ 11 summary line
      const text11 =
        `11. รายการอุปกรณ์ที่ได้รับความเสียหาย ${totalItems} รายการ ` +
        `คิดเป็นค่าเสียหาย จำนวนเงินทั้งสิ้น ${fmt(totalAmount)} บาท โดยมีรายละเอียดดังนี้`;
      
      const linesIndent11 = pdf.splitTextToSize(text11, cW - 30);
      const line1_11 = linesIndent11[0] || '';
      const remainingText11 = text11.substring(line1_11.length).trim();
      const remainingLines11 = pdf.splitTextToSize(remainingText11, cW);

      pdf.text(line1_11, mL + 30, y);
      if (remainingLines11.length > 0) {
        y += lh;
        pdf.text(remainingLines11, mL, y);
        y += remainingLines11.length * lh + 2;
      } else {
        y += lh + 2;
      }

      // Table column widths (total = 160 mm = cW)
      const colWidths = [18, 83, 18, 14, 27];

      // Helper: render autoTable and return finalY
      const drawTable = (startY: number, bodyRows: string[][]): number => {
        autoTable(pdf, {
          startY,
          head: [['รายการ', 'ชื่อพัสดุ', 'จำนวน', 'หน่วย', 'ราคา']],
          body: bodyRows,
          margin: { left: mL, right: mR },
          tableWidth: cW,
          styles: {
            font: 'THSarabunNew',
            fontStyle: 'normal',
            fontSize: 14,
            cellPadding: { top: 2, bottom: 2, left: 3, right: 3 },
            lineColor: [0, 0, 0],
            lineWidth: 0.3,
            textColor: [0, 0, 0],
            valign: 'middle',
          },
          headStyles: {
            font: 'THSarabunNew',
            fontStyle: 'bold',
            fontSize: 14,
            fillColor: [242, 242, 242],
            textColor: [0, 0, 0],
            halign: 'center',
          },
          columnStyles: {
            0: { halign: 'center', cellWidth: colWidths[0], cellPadding: { top: 2, bottom: 2, left: 1, right: 1 } },
            1: { halign: 'left',   cellWidth: colWidths[1] },
            2: { halign: 'right',  cellWidth: colWidths[2] },
            3: { halign: 'center', cellWidth: colWidths[3] },
            4: { halign: 'right',  cellWidth: colWidths[4] },
          },
          theme: 'grid',
        });
        return (pdf as any).lastAutoTable.finalY as number;
      };

      // ---- 11.1 รื้อถอน - ติดตั้งใหม่ ----
      pdf.text(
        `11.1 รื้อถอน - ติดตั้งใหม่ ${damagedItems.length} รายการ เป็นจำนวนเงินทั้งสิ้น ${fmt(damagedTotal)} บาท`,
        mL + 35, y,
      );
      y += 4;

      const damagedRows: string[][] = damagedItems.map((item, idx) => [
        String(idx + 1),
        item.material.name,
        item.quantity.toLocaleString(),
        item.material.unit,
        fmt(item.totalPrice),
      ]);
      while (damagedRows.length < 5) damagedRows.push(['', '', '', '', '']);

      y = drawTable(y, damagedRows) + 5;
      y += lh; // เคาะบรรทัด 1 ครั้ง

      // ---- 11.2 แผนกซ่อมแซม ----
      setF('normal', 16);
      pdf.text(
        `11.2 แผนกซ่อมแซม ${reusableItems.length} รายการ เป็นจำนวนเงินทั้งสิ้น ${fmt(reusableTotal)} บาท`,
        mL + 35, y,
      );
      y += 4;

      const reusableRows: string[][] = reusableItems.map((item, idx) => [
        String(idx + 1),
        item.material.name,
        item.quantity.toLocaleString(),
        item.material.unit,
        fmt(item.totalPrice),
      ]);
      while (reusableRows.length < 5) reusableRows.push(['', '', '', '', '']);

      y = drawTable(y, reusableRows) + 5;
      y += lh; // เคาะบรรทัด 1 ครั้ง

      // ---- วรรคปิด ----
      setF('normal', 16);
      const closingText =
        `จึงเรียนมาเพื่อพิจารณาอนุมัติให้ดำเนินการเบิกอุปกรณ์ไปซ่อมแซมตามรายการดังกล่าว ` +
        `พร้อมทั้งเป็นการเรียกเก็บเงินค่าเสียหายจากผู้กระทำละเมิด เป็นจำนวนเงินทั้งสิ้น ${fmt(totalAmount)} บาท`;
      
      const closingLinesIndent = pdf.splitTextToSize(closingText, cW - 30);
      const closingLine1 = closingLinesIndent[0] || '';
      const remainingClosingText = closingText.substring(closingLine1.length).trim();
      const remainingClosingLines = pdf.splitTextToSize(remainingClosingText, cW);

      pdf.text([closingLine1, ''], mL + 30, y, { maxWidth: cW - 30, align: 'justify' });
      if (remainingClosingLines.length > 0) {
        y += lh;
        pdf.text(remainingClosingLines, mL, y, { maxWidth: cW, align: 'justify' });
        y += remainingClosingLines.length * lh + 18;
      } else {
        y += lh + 18;
      }

      // ---- ลายเซ็น ----
      const sigX = 100; // 10 cm จากขอบกระดาษซ้ายสุด
      pdf.text('(...........................................................)', sigX, y);
      y += lh;
      const lineWidth = pdf.getTextWidth('(...........................................................)');
      const sigCenterX = sigX + lineWidth * 0.5;
      pdf.text('ตำแหน่ง', sigCenterX, y, { align: 'center' });

      // ---- Footer page 2 ----
      setF('normal', 14);
      pdf.text('หน่วยงาน', mL, pageH - mB + 5);
      pdf.text('โทร. ...........................................................', mL, pageH - mB + 12);

      // ---- Page numbers for page 3+ (if tables overflow to next page) ----
      const totalPages = pdf.getNumberOfPages();
      for (let i = 3; i <= totalPages; i++) {
        pdf.setPage(i);
        setF('normal', 14);
        pdf.text(`- ${i} -`, pageW / 2, 12, { align: 'center' });
      }

      pdf.save('ค่าละเมิด1234_PI InnoTech.pdf');
    } catch (error) {
      console.error('Error generating PDF:', error);
      alert('เกิดข้อผิดพลาดในการสร้างไฟล์ PDF');
    } finally {
      setExporting(false);
    }
  };

  if (loading) {
    return (
      <div className="flex flex-col items-center justify-center min-h-screen bg-slate-50">
        <Loader2 className="w-12 h-12 text-blue-600 animate-spin mb-4" />
        <p className="text-slate-600 font-medium">กำลังโหลดข้อมูลพัสดุ...</p>
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-slate-50 pb-20">
      {/* Header */}
      <header className="bg-white border-b border-slate-200 sticky top-0 z-10">
        <div className="max-w-6xl mx-auto px-4 py-3 flex flex-col sm:flex-row items-start sm:items-center justify-between gap-3">
          <div className="flex items-start gap-3">
            <div className="bg-blue-600 p-2.5 rounded-xl shadow-lg shadow-blue-100 mt-1">
              <Settings2 className="w-6 h-6 text-white" />
            </div>
            <div className="space-y-1">
              <h1 className="text-lg md:text-xl font-bold text-slate-900 leading-tight">
                ค่าละเมิด1234
              </h1>
              <p className="text-xs md:text-sm text-slate-500 font-medium">
                ราคาพัสดุแบบสำเร็จรูป สำหรับการประเมินค่าเสียหายที่เกิดกับระบบจำหน่ายและระบบสายส่ง : กฟต.3
              </p>
              
              {/* Dropdown list for selecting sheet/version directly below "ราคาพัสดุแบบสำเร็จรูป" */}
              <div className="flex items-center gap-2 pt-1">
                <span className="text-xs font-semibold text-slate-600 shrink-0">
                  เลือกฉบับราคา:
                </span>
                <div className="relative inline-flex items-center">
                  <select
                    id="sheet-select"
                    value={selectedSheet}
                    onChange={(e) => {
                      setSelectedSheet(e.target.value);
                    }}
                    className="text-xs font-semibold bg-blue-50/80 hover:bg-blue-100/80 text-blue-800 border border-blue-200 hover:border-blue-300 rounded-lg px-2.5 py-1 pr-7 outline-none focus:ring-2 focus:ring-blue-500 transition-all cursor-pointer shadow-2xs appearance-none"
                  >
                    {sheetOptions.map((sheet) => (
                      <option key={sheet} value={sheet}>
                        {sheet}
                      </option>
                    ))}
                  </select>
                  <ChevronDown className="w-3.5 h-3.5 text-blue-600 absolute right-2 pointer-events-none" />
                </div>
              </div>
            </div>
          </div>
          <div className="flex items-center gap-4 self-end sm:self-center">
            <div className="text-right hidden sm:block">
              <p className="text-[10px] text-slate-400 uppercase tracking-widest font-bold">Total Assessment</p>
              <p className="text-xl font-black text-blue-600">฿{totalAmount.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })}</p>
            </div>
          </div>
        </div>
      </header>

      <main className="max-w-6xl mx-auto px-4 py-8">
        {error && (
          <div className="mb-6 p-6 bg-red-50 border border-red-200 rounded-2xl flex flex-col items-center gap-4 text-center">
            <div className="flex items-center gap-3 text-red-700">
              <AlertCircle className="w-6 h-6 flex-shrink-0" />
              <p className="font-semibold">{error}</p>
            </div>
            <button 
              onClick={fetchData}
              className="flex items-center gap-2 px-6 py-2 bg-red-600 text-white font-bold rounded-xl hover:bg-red-700 transition-all shadow-md"
            >
              <RefreshCw className="w-4 h-4" />
              ลองใหม่อีกครั้ง
            </button>
          </div>
        )}

        <div className="grid grid-cols-1 lg:grid-cols-2 gap-8">
          {/* Left Column: Single Page Form */}
          <div className="space-y-6">
            <div className="bg-white rounded-2xl shadow-sm border border-slate-200 overflow-hidden p-6 space-y-8">
              {/* Step 1: Search */}
              <div className="space-y-4">
                <div className="flex items-center gap-2 text-blue-600">
                  <h2 className="font-bold">STEP 1 : พิมพ์ชื่อพัสดุหรือรหัสพัสดุ กฟภ.</h2>
                </div>
                
                <div className="space-y-2">
                  <div className="relative">
                    <Search className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400 w-5 h-5" />
                    <input 
                      type="text"
                      placeholder="พิมพ์ชื่อหรือรหัสพัสดุ..."
                      className="w-full pl-10 pr-4 py-3 bg-slate-50 border border-slate-200 rounded-xl focus:ring-2 focus:ring-blue-500 outline-none transition-all"
                      value={searchQuery}
                      onChange={(e) => {
                        setSearchQuery(e.target.value);
                        if (!e.target.value) setSelectedMaterial(null);
                      }}
                    />
                  </div>

                  {selectedMaterial ? (
                    <div className="p-4 bg-blue-50 border border-blue-100 rounded-xl flex items-center justify-between">
                      <div>
                        <p className="font-bold text-blue-900">{selectedMaterial.name}</p>
                        <p className="text-sm text-blue-700">
                          รหัส: {selectedMaterial.id} | 
                          ชำรุด: ฿{selectedMaterial.priceDamaged.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })} | 
                          นำกลับมาใช้ใหม่: ฿{selectedMaterial.priceReusable.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
                        </p>
                      </div>
                      <button 
                        onClick={() => {
                          setSelectedMaterial(null);
                          setSearchQuery('');
                        }}
                        className="p-2 hover:bg-blue-100 rounded-full text-blue-600 transition-colors"
                      >
                        <RefreshCw className="w-4 h-4" />
                      </button>
                    </div>
                  ) : (
                    <div className="space-y-1 max-h-48 overflow-y-auto custom-scrollbar">
                      {filteredMaterials.map((m) => (
                        <button
                          key={m.id}
                          onClick={() => {
                            setSelectedMaterial(m);
                            setSearchQuery(m.name);
                          }}
                          className="w-full text-left p-3 hover:bg-slate-50 border border-transparent hover:border-slate-200 rounded-xl transition-all flex items-center justify-between group"
                        >
                          <div>
                            <p className="font-medium text-slate-900">{m.name}</p>
                            <p className="text-xs text-slate-500">รหัส: {m.id}</p>
                          </div>
                          <ChevronRight className="w-4 h-4 text-slate-300 group-hover:text-blue-500 transition-colors" />
                        </button>
                      ))}
                    </div>
                  )}
                </div>
              </div>

              {/* Step 2 & 3: Quantity and Status */}
              <div className={cn("grid grid-cols-1 md:grid-cols-2 gap-8 pt-6 border-t border-slate-100 transition-opacity", !selectedMaterial && "opacity-50 pointer-events-none")}>
                {/* Quantity */}
                <div className="space-y-4">
                  <div className="flex items-center gap-2 text-blue-600">
                    <h2 className="font-bold">STEP 2 : ระบุจำนวน ({selectedMaterial?.unit || '-'})</h2>
                  </div>
                  <div className="space-y-4">
                    <input 
                      type="text"
                      inputMode="numeric"
                      pattern="[0-9]*"
                      disabled={!selectedMaterial}
                      className="w-full px-4 py-3 bg-slate-50 border border-slate-200 rounded-xl focus:ring-2 focus:ring-blue-500 outline-none text-2xl font-bold text-center"
                      value={quantityInput}
                      onChange={(e) => {
                        const rawValue = e.target.value;
                        const cleaned = rawValue.replace(/[^0-9]/g, '');
                        if (cleaned === '') {
                          setQuantityInput('');
                          setQuantity(1);
                          return;
                        }
                        const parsed = parseInt(cleaned, 10);
                        if (!isNaN(parsed)) {
                          setQuantityInput(parsed.toLocaleString());
                          setQuantity(parsed);
                        }
                      }}
                    />
                    <div className="grid grid-cols-4 gap-2">
                      {[1, 2, 5, 10].map(n => (
                        <button 
                          key={n}
                          disabled={!selectedMaterial}
                          onClick={() => setQuantity(n)}
                          className="py-2 bg-white border border-slate-200 rounded-lg text-slate-600 hover:border-blue-500 hover:text-blue-600 transition-all text-sm font-medium"
                        >
                          {n}
                        </button>
                      ))}
                    </div>
                  </div>
                </div>

                {/* Status Selection */}
                <div className="space-y-4">
                  <div className="flex items-center gap-2 text-blue-600">
                    <h2 className="font-bold">STEP 3 : เลือกสถานะเพื่อบันทึก</h2>
                  </div>
                  <div className="grid grid-cols-1 gap-3">
                    <button
                      disabled={!selectedMaterial}
                      onClick={() => handleStatusSelect('damaged')}
                      className="flex items-center justify-between p-4 rounded-xl border-2 border-slate-100 bg-white text-slate-600 hover:bg-red-50 hover:border-red-500 hover:text-red-700 transition-all group"
                    >
                      <div className="flex items-center gap-3">
                        <AlertCircle className="w-6 h-6 text-slate-300 group-hover:text-red-500 transition-colors" />
                        <span className="font-bold">ชำรุด</span>
                      </div>
                      <span className="text-xs font-bold bg-red-100 text-red-600 px-2 py-1 rounded">฿{((selectedMaterial?.priceDamaged || 0) * quantity).toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })}</span>
                    </button>
                    <button
                      disabled={!selectedMaterial}
                      onClick={() => handleStatusSelect('reusable')}
                      className="flex items-center justify-between p-4 rounded-xl border-2 border-slate-100 bg-white text-slate-600 hover:bg-green-50 hover:border-green-500 hover:text-green-700 transition-all group"
                    >
                      <div className="flex items-center gap-3">
                        <RefreshCw className="w-6 h-6 text-slate-300 group-hover:text-green-500 transition-colors" />
                        <span className="font-bold">นำกลับมาใช้ใหม่</span>
                      </div>
                      <span className="text-xs font-bold bg-green-100 text-green-600 px-2 py-1 rounded">฿{((selectedMaterial?.priceReusable || 0) * quantity).toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })}</span>
                    </button>
                  </div>
                </div>
              </div>

              {!selectedMaterial && (
                <div className="text-center py-4 text-slate-400 text-sm italic">
                  * กรุณาเลือกพัสดุในขั้นตอนที่ 1 ก่อนระบุจำนวนและสถานะ
                </div>
              )}
            </div>
          </div>

          {/* Right Column: Assessment Summary */}
          <div className="space-y-6">
            <div className="bg-white rounded-2xl shadow-sm border border-slate-200 p-6 flex flex-col h-full max-h-[700px]">
              <div className="flex items-center justify-between mb-6">
                <h2 className="text-lg font-bold text-slate-900 flex items-center gap-2">
                  <FileText className="w-5 h-5 text-blue-600" />
                  รายการประเมิน
                </h2>
                <div className="flex items-center gap-2">
                  {items.length > 0 && (
                    <button
                      onClick={() => setShowClearConfirm(true)}
                      className="text-xs text-red-600 hover:text-white hover:bg-red-600 font-medium px-2.5 py-1 rounded-lg border border-red-200 hover:border-red-600 transition-all flex items-center gap-1 cursor-pointer active:scale-95"
                    >
                      <Trash2 className="w-3 h-3" />
                      ล้างรายการทั้งหมด
                    </button>
                  )}
                  <span className="bg-blue-100 text-blue-700 text-xs font-bold px-2 py-1 rounded-full">
                    {items.length} รายการ
                  </span>
                </div>
              </div>

              <div className="flex-1 overflow-y-auto space-y-3 pr-2 custom-scrollbar">
                {items.length === 0 ? (
                  <div className="flex flex-col items-center justify-center py-12 text-slate-400">
                    <Package className="w-12 h-12 mb-2 opacity-20" />
                    <p className="text-sm italic">ยังไม่มีรายการประเมิน</p>
                  </div>
                ) : (
                  items.map((item, index) => (
                    <motion.div 
                      initial={{ opacity: 0, y: 10 }}
                      animate={{ opacity: 1, y: 0 }}
                      key={index}
                      className="p-3 bg-slate-50 rounded-xl border border-slate-100 group relative"
                    >
                      <button 
                        onClick={() => removeItem(index)}
                        className="absolute -top-2 -right-2 w-6 h-6 bg-white border border-slate-200 rounded-full flex items-center justify-center text-red-500 shadow-sm opacity-0 group-hover:opacity-100 transition-opacity"
                      >
                        <Trash2 className="w-3 h-3" />
                      </button>
                      <div className="flex justify-between items-start mb-1">
                        <p className="font-bold text-sm text-slate-900 line-clamp-1">{item.material.name}</p>
                        <span className={cn(
                          "text-[10px] font-bold px-1.5 py-0.5 rounded uppercase",
                          item.status === 'damaged' ? "bg-red-100 text-red-700" : "bg-green-100 text-green-700"
                        )}>
                          {item.status === 'damaged' ? 'ชำรุด' : 'นำกลับมาใช้ใหม่'}
                        </span>
                      </div>
                      <div className="flex justify-between items-end">
                        <p className="text-xs text-slate-500">
                          {item.quantity.toLocaleString()} {item.material.unit} x ฿{(item.status === 'damaged' ? item.material.priceDamaged : item.material.priceReusable).toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
                        </p>
                        <p className="font-bold text-blue-600">฿{item.totalPrice.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })}</p>
                      </div>
                    </motion.div>
                  ))
                )}
              </div>

              <div className="mt-6 pt-6 border-t border-slate-100 space-y-4">
                <div className="flex justify-between items-center">
                  <span className="text-slate-500 font-medium">รวมทั้งหมด</span>
                  <span className="text-2xl font-black text-slate-900">฿{totalAmount.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })}</span>
                </div>
                
                <div className="space-y-4">
                  <div className="flex items-center gap-2 text-blue-600">
                    <h2 className="font-bold">STEP 4 : ออกรายการประเมินค่าเสียหาย</h2>
                  </div>
                  <div className="flex gap-3">
                    <button 
                      disabled={items.length === 0}
                      onClick={exportWord}
                      className="flex-1 flex flex-col items-center justify-center gap-1 py-4 bg-[#EF0107] text-white font-bold rounded-xl hover:bg-[#DB0007] transition-all disabled:opacity-50 disabled:cursor-not-allowed shadow-lg shadow-red-100"
                    >
                      <div className="flex items-center gap-2">
                        <Download className="w-5 h-5" />
                        Export Word
                      </div>
                      <span className="text-[10px] font-normal opacity-90">(For Notebook & PC)</span>
                    </button>
                    <button 
                      disabled={items.length === 0 || exporting}
                      onClick={exportPDF}
                      className="flex-1 flex flex-col items-center justify-center gap-1 py-4 bg-slate-800 text-white font-bold rounded-xl hover:bg-slate-900 transition-all disabled:opacity-50 disabled:cursor-not-allowed shadow-lg shadow-slate-100"
                    >
                      <div className="flex items-center gap-2">
                        {exporting ? (
                          <Loader2 className="w-5 h-5 animate-spin" />
                        ) : (
                          <FileText className="w-5 h-5" />
                        )}
                        {exporting ? 'Exporting...' : 'Export PDF'}
                      </div>
                      <span className="text-[10px] font-normal opacity-90">(For Tablet & Smart Phone)</span>
                    </button>
                  </div>
                </div>
              </div>
            </div>
          </div>
        </div>
      </main>

      {/* PDF is now generated programmatically via jsPDF + autoTable — no hidden template needed */}
      <div style={{ display: 'none' }}>
        <div>
          {/* Page 1 */}
          <div style={{ minHeight: '277mm', position: 'relative' }}>
            <img 
              src="https://img1.pic.in.th/images/PEA-02-Thai-Logo.md.jpg" 
              alt="Logo" 
              style={{ width: '41.5mm', height: '34.9mm', marginBottom: '3mm' }}
              referrerPolicy="no-referrer"
            />
            
            <div style={{ display: 'flex', marginBottom: '3mm' }}>
              <div style={{ width: '80mm' }}>จาก</div>
              <div style={{ width: '80mm', marginLeft: '0mm' }}>ถึง</div>
            </div>
            <div style={{ display: 'flex', marginBottom: '5mm' }}>
              <div style={{ width: '80mm' }}>เลขที่</div>
              <div style={{ width: '80mm', marginLeft: '0mm' }}>วันที่</div>
            </div>

            <div style={{ marginBottom: '3mm' }}>เรื่อง&nbsp;&nbsp;&nbsp;การประเมินค่าเสียหายที่เกิดขึ้นกับระบบจำหน่าย</div>
            <div style={{ marginBottom: '5mm' }}>เรียน</div>

            <div style={{ 
              textIndent: '25mm', 
              textAlign: 'justify', 
              textJustify: 'inter-character', 
              wordBreak: 'break-word',
              marginBottom: '2mm',
              lineHeight: '1.5',
              letterSpacing: '0.2px',
              textRendering: 'optimizeLegibility'
            }}>
              ตามที่ได้ดำเนินการตรวจสอบและประเมินราคาค่าเสียหายเพื่อเรียกร้องจากผู้กระทำละเมิด โดยมีรายละเอียด ดังนี้
            </div>

            <div style={{ marginLeft: '25mm' }}>
              <div style={{ marginBottom: '2mm' }}>1. เหตุเกิดเมื่อ ..................................................................................................</div>
              <div style={{ marginBottom: '2mm' }}>2. สถานที่เกิดเหตุ ............................................................................................</div>
              <div style={{ marginBottom: '2mm' }}>3. หมายเลขทะเบียน ........................................................................................</div>
              <div style={{ marginBottom: '2mm' }}>4. ชื่อผู้ขับขี่ .....................................................................................................</div>
              <div style={{ marginBottom: '2mm', marginLeft: '5mm' }}>บัตรประชาชนเลขที่ ...................................................................</div>
              <div style={{ marginBottom: '2mm' }}>5. ที่อยู่ตามบัตร ...............................................................................................</div>
              <div style={{ display: 'flex', marginBottom: '2mm', marginLeft: '-25mm' }}>
                <div style={{ width: '100mm' }}>............................................................................................</div>
                <div style={{ width: '50mm' }}>เบอร์โทรศัพท์ ......................</div>
              </div>
              <div style={{ display: 'flex', marginBottom: '2mm' }}>
                <div style={{ width: '75mm' }}>6. ชื่อ/บริษัท เจ้าของรถยนต์ .....................</div>
                <div style={{ width: '50mm' }}>เบอร์โทรศัพท์ ......................</div>
              </div>
              <div style={{ display: 'flex', marginBottom: '2mm' }}>
                <div style={{ width: '75mm' }}>7. ชื่อ/บริษัท ประกันภัย ............................</div>
                <div style={{ width: '50mm' }}>เบอร์โทรศัพท์ ......................</div>
              </div>
              <div style={{ marginBottom: '2mm' }}>8. ผู้ลงนามในหนังสือรับสภาพหนี้</div>
              <div style={{ marginBottom: '2mm', marginLeft: '5mm' }}>[  ] ผู้ขับขี่    [  ] เจ้าของรถยนต์    [  ] ไม่ยินยอม</div>
              <div style={{ marginBottom: '2mm' }}>9. การแจ้งความร้องทุกข์กับเจ้าหน้าที่ตำรวจ</div>
              <div style={{ marginBottom: '2mm', marginLeft: '5mm' }}>[  ] แจ้งเป็นหลักฐาน    [  ] แจ้งความเป็นคดี เนื่องจาก .....................................</div>
              <div style={{ marginBottom: '2mm' }}>10. กรณีรถยนต์เกี่ยวสายสื่อสารทำให้เกิดความเสียหายกับระบบจำหน่าย</div>
              <div style={{ marginBottom: '2mm', marginLeft: '5mm' }}>ชื่อ/บริษัท เจ้าของสายสื่อสาร .............................. ความสูง ........................</div>
            </div>
          </div>

          {/* Page 2 */}
          <div style={{ minHeight: '242mm', position: 'relative', paddingTop: '20mm' }}>
            <div style={{ position: 'absolute', top: '20mm', left: '50%', transform: 'translateX(-50%)', fontSize: '12pt' }}>&nbsp;</div>
            
            <div style={{ textIndent: '25mm', textAlign: 'justify', marginBottom: '2mm' }}>
              11. รายการอุปกรณ์ที่ได้รับความเสียหาย {groupItems(items).length} รายการ คิดเป็นค่าเสียหาย จำนวนเงินทั้งสิ้น {totalAmount.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })} บาท โดยมีรายละเอียดดังนี้
            </div>

            <div style={{ marginBottom: '5mm', marginLeft: '30mm' }}>
              11.1 รื้อถอน - ติดตั้งใหม่ {groupItems(items.filter(i => i.status === 'damaged')).length} รายการ เป็นจำนวนเงินทั้งสิ้น {groupItems(items.filter(i => i.status === 'damaged')).reduce((sum, i) => sum + i.totalPrice, 0).toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })} บาท
            </div>

            <table style={{ width: '100%', borderCollapse: 'collapse', marginBottom: '5mm', fontSize: '10pt' }}>
              <thead>
                <tr style={{ backgroundColor: '#f2f2f2' }}>
                  <th style={{ border: '1px solid black', padding: '10px 5px', width: '10%', textAlign: 'center', verticalAlign: 'middle', lineHeight: '1.2' }}>รายการ</th>
                  <th style={{ border: '1px solid black', padding: '10px 5px', width: '50%', textAlign: 'center', verticalAlign: 'middle', lineHeight: '1.2' }}>ชื่อพัสดุ</th>
                  <th style={{ border: '1px solid black', padding: '10px 5px', width: '10%', textAlign: 'center', verticalAlign: 'middle', lineHeight: '1.2' }}>จำนวน</th>
                  <th style={{ border: '1px solid black', padding: '10px 5px', width: '10%', textAlign: 'center', verticalAlign: 'middle', lineHeight: '1.2' }}>หน่วย</th>
                  <th style={{ border: '1px solid black', padding: '10px 5px', width: '20%', textAlign: 'center', verticalAlign: 'middle', lineHeight: '1.2' }}>ราคา</th>
                </tr>
              </thead>
              <tbody>
                {groupItems(items.filter(i => i.status === 'damaged')).map((item, idx) => (
                  <tr key={idx}>
                    <td style={{ border: '1px solid black', padding: '8px 5px', textAlign: 'center' }}>{idx + 1}</td>
                    <td style={{ border: '1px solid black', padding: '8px 5px' }}>{item.material.name}</td>
                    <td style={{ border: '1px solid black', padding: '8px 5px', textAlign: 'center' }}>{item.quantity.toLocaleString()}</td>
                    <td style={{ border: '1px solid black', padding: '8px 5px', textAlign: 'center' }}>{item.material.unit}</td>
                    <td style={{ border: '1px solid black', padding: '8px 5px', textAlign: 'right' }}>{item.totalPrice.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })}</td>
                  </tr>
                ))}
                {Array.from({ length: Math.max(0, 5 - groupItems(items.filter(i => i.status === 'damaged')).length) }).map((_, idx) => (
                  <tr key={`empty-${idx}`}>
                    <td style={{ border: '1px solid black', padding: '8px 5px', height: '1.5em' }}></td>
                    <td style={{ border: '1px solid black', padding: '8px 5px' }}></td>
                    <td style={{ border: '1px solid black', padding: '8px 5px' }}></td>
                    <td style={{ border: '1px solid black', padding: '8px 5px' }}></td>
                    <td style={{ border: '1px solid black', padding: '8px 5px' }}></td>
                  </tr>
                ))}
              </tbody>
            </table>

            <div style={{ marginBottom: '5mm', marginTop: '5mm', marginLeft: '30mm' }}>
              11.2 แผนกซ่อมแซม {groupItems(items.filter(i => i.status === 'reusable')).length} รายการ เป็นจำนวนเงินทั้งสิ้น {groupItems(items.filter(i => i.status === 'reusable')).reduce((sum, i) => sum + i.totalPrice, 0).toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })} บาท
            </div>

            <table style={{ width: '100%', borderCollapse: 'collapse', marginBottom: '5mm', fontSize: '10pt' }}>
              <thead>
                <tr style={{ backgroundColor: '#f2f2f2' }}>
                  <th style={{ border: '1px solid black', padding: '10px 5px', width: '10%', textAlign: 'center', verticalAlign: 'middle', lineHeight: '1.2' }}>รายการ</th>
                  <th style={{ border: '1px solid black', padding: '10px 5px', width: '50%', textAlign: 'center', verticalAlign: 'middle', lineHeight: '1.2' }}>ชื่อพัสดุ</th>
                  <th style={{ border: '1px solid black', padding: '10px 5px', width: '10%', textAlign: 'center', verticalAlign: 'middle', lineHeight: '1.2' }}>จำนวน</th>
                  <th style={{ border: '1px solid black', padding: '10px 5px', width: '10%', textAlign: 'center', verticalAlign: 'middle', lineHeight: '1.2' }}>หน่วย</th>
                  <th style={{ border: '1px solid black', padding: '10px 5px', width: '20%', textAlign: 'center', verticalAlign: 'middle', lineHeight: '1.2' }}>ราคา</th>
                </tr>
              </thead>
              <tbody>
                {groupItems(items.filter(i => i.status === 'reusable')).map((item, idx) => (
                  <tr key={idx}>
                    <td style={{ border: '1px solid black', padding: '8px 5px', textAlign: 'center' }}>{idx + 1}</td>
                    <td style={{ border: '1px solid black', padding: '8px 5px' }}>{item.material.name}</td>
                    <td style={{ border: '1px solid black', padding: '8px 5px', textAlign: 'center' }}>{item.quantity.toLocaleString()}</td>
                    <td style={{ border: '1px solid black', padding: '8px 5px', textAlign: 'center' }}>{item.material.unit}</td>
                    <td style={{ border: '1px solid black', padding: '8px 5px', textAlign: 'right' }}>{item.totalPrice.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })}</td>
                  </tr>
                ))}
                {Array.from({ length: Math.max(0, 5 - groupItems(items.filter(i => i.status === 'reusable')).length) }).map((_, idx) => (
                  <tr key={`empty-re-${idx}`}>
                    <td style={{ border: '1px solid black', padding: '8px 5px', height: '1.5em' }}></td>
                    <td style={{ border: '1px solid black', padding: '8px 5px' }}></td>
                    <td style={{ border: '1px solid black', padding: '8px 5px' }}></td>
                    <td style={{ border: '1px solid black', padding: '8px 5px' }}></td>
                    <td style={{ border: '1px solid black', padding: '8px 5px' }}></td>
                  </tr>
                ))}
              </tbody>
            </table>

            <div style={{ textAlign: 'justify', textJustify: 'inter-character', marginBottom: '30mm', marginTop: '5mm', lineHeight: '1.6', letterSpacing: '0.25px' }}>
              &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;จึงเรียนมาเพื่อพิจารณาอนุมัติให้ดำเนินการเบิกอุปกรณ์ไปซ่อมแซมตามรายการดังกล่าว พร้อมทั้งเป็นการเรียกเก็บเงินค่าเสียหายจากผู้กระทำละเมิด เป็นจำนวนเงินทั้งสิ้น {totalAmount.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })} บาท
            </div>

            <div style={{ textAlign: 'center', marginLeft: '80mm' }}>
              <div style={{ marginBottom: '5mm' }}>(...........................................................)</div>
              <div>ตำแหน่ง</div>
            </div>

            <div style={{ position: 'absolute', bottom: '30', left: '0', fontSize: '10pt' }}>
              <div>หน่วยงาน</div>
              <div>โทร. ...........................................................</div>
            </div>
          </div>
        </div>
      </div>

      {/* Footer */}
      <footer className="max-w-6xl mx-auto px-4 py-12 text-center space-y-2">
        <p className="text-slate-400 text-xs font-medium">
          © 2026 ค่าละเมิด1234 - <span className="text-blue-500/80">PI InnoTech</span>
        </p>
        <p className="text-slate-300 text-[9px] font-bold tracking-[0.25em] uppercase">
          PROVINCIAL ELECTRICITY AUTHORITY
        </p>
      </footer>

      {/* Mobile Bottom Bar for Total */}
      <div className="sm:hidden fixed bottom-0 left-0 right-0 bg-white border-t border-slate-200 p-4 flex items-center justify-between z-20">
        <div>
          <p className="text-[10px] text-slate-500 uppercase font-bold">Total Assessment</p>
          <p className="text-xl font-black text-blue-600">฿{totalAmount.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })}</p>
        </div>
        <div className="flex gap-2">
          <button 
            disabled={items.length === 0}
            onClick={exportWord}
            className="bg-[#EF0107] text-white px-4 py-3 rounded-xl font-bold text-xs flex items-center gap-2 hover:bg-[#DB0007] transition-colors"
          >
            <Download className="w-4 h-4" />
            Word
          </button>
          <button 
            disabled={items.length === 0 || exporting}
            onClick={exportPDF}
            className="bg-slate-800 text-white px-4 py-3 rounded-xl font-bold text-xs flex items-center gap-2 hover:bg-slate-900 transition-colors"
          >
            {exporting ? (
              <Loader2 className="w-4 h-4 animate-spin" />
            ) : (
              <FileText className="w-4 h-4" />
            )}
            {exporting ? '...' : 'PDF'}
          </button>
        </div>
      </div>

      {/* Confirm Clear Modal */}
      <AnimatePresence>
        {showClearConfirm && (
          <div className="fixed inset-0 z-50 flex items-center justify-center p-4">
            {/* Backdrop */}
            <motion.div
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
              exit={{ opacity: 0 }}
              onClick={() => setShowClearConfirm(false)}
              className="absolute inset-0 bg-slate-900/60 backdrop-blur-xs"
            />

            {/* Modal Content Card */}
            <motion.div
              initial={{ opacity: 0, scale: 0.95, y: 16 }}
              animate={{ opacity: 1, scale: 1, y: 0 }}
              exit={{ opacity: 0, scale: 0.95, y: 16 }}
              transition={{ duration: 0.2, ease: "easeOut" }}
              className="relative w-full max-w-md bg-white rounded-2xl shadow-xl border border-slate-100 p-6 space-y-6 text-center z-10 overflow-hidden"
            >
              {/* Decorative Accent Header */}
              <div className="absolute top-0 inset-x-0 h-1.5 bg-gradient-to-r from-red-500 to-orange-500" />
              
              <div className="flex flex-col items-center gap-4 pt-2">
                <div className="w-12 h-12 rounded-full bg-red-50 flex items-center justify-center text-red-500 border border-red-100">
                  <Trash2 className="w-6 h-6" />
                </div>
                <div className="space-y-1.5">
                  <h3 className="text-lg font-bold text-slate-900">
                    ยืนยันการล้างรายการทั้งหมด?
                  </h3>
                  <p className="text-sm text-slate-500 leading-relaxed px-2">
                    คุณแน่ใจหรือไม่ว่าต้องการลบรายการประเมินค่าเสียหายทั้งหมด? การดำเนินการนี้จะลบข้อมูลปัจจุบันของคุณทั้งหมดและไม่สามารถกู้คืนได้
                  </p>
                </div>
              </div>

              <div className="flex gap-3 justify-center">
                <button
                  type="button"
                  onClick={() => setShowClearConfirm(false)}
                  className="flex-1 py-2.5 px-4 rounded-xl border border-slate-200 text-slate-700 hover:bg-slate-50 active:scale-98 transition-all font-medium text-sm cursor-pointer"
                >
                  ยกเลิก
                </button>
                <button
                  type="button"
                  onClick={() => {
                    setItems([]);
                    setShowClearConfirm(false);
                  }}
                  className="flex-1 py-2.5 px-4 rounded-xl bg-red-600 hover:bg-red-700 text-white font-semibold text-sm shadow-sm hover:shadow active:scale-98 transition-all cursor-pointer"
                >
                  ล้างรายการทั้งหมด
                </button>
              </div>
            </motion.div>
          </div>
        )}
      </AnimatePresence>
    </div>
  );
}
