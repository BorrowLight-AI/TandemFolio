// Modified by TandemFolio contributors: native streamed-find status copy.
import { defineStrings } from '@genoffice/i18n'

/** Status shown when bounded native Find cannot scan the complete workbook. */
export const nativeFindStrings = defineStrings({
  zh: { appFindScanTruncated: '查找已扫描 {cells} 个单元格；工作簿过大，结果可能不完整。' },
  en: { appFindScanTruncated: 'Find scanned {cells} cells; this large workbook may have more matches.' },
  ja: { appFindScanTruncated: '検索で {cells} セルを走査しました。大きなブックのため結果が不完全な場合があります。' },
  ko: { appFindScanTruncated: '찾기에서 {cells}개 셀을 검사했습니다. 큰 통합 문서에는 결과가 더 있을 수 있습니다.' },
  fr: { appFindScanTruncated: 'La recherche a analysé {cells} cellules ; ce grand classeur peut contenir d’autres résultats.' },
  de: { appFindScanTruncated: 'Die Suche hat {cells} Zellen geprüft; diese große Arbeitsmappe kann weitere Treffer enthalten.' },
  es: { appFindScanTruncated: 'Buscar examinó {cells} celdas; este libro grande puede contener más coincidencias.' },
  th: { appFindScanTruncated: 'ค้นหาตรวจสอบแล้ว {cells} เซลล์ สมุดงานขนาดใหญ่นี้อาจมีผลลัพธ์เพิ่มเติม' },
  id: { appFindScanTruncated: 'Pencarian memindai {cells} sel; buku kerja besar ini mungkin memiliki hasil lain.' },
  ru: { appFindScanTruncated: 'Поиск проверил {cells} ячеек; в этой большой книге могут быть другие совпадения.' },
  ar: { appFindScanTruncated: 'فحص البحث {cells} خلية؛ قد يحتوي هذا المصنف الكبير على نتائج أخرى.' },
  pt: { appFindScanTruncated: 'A pesquisa examinou {cells} células; esta pasta de trabalho grande pode ter mais resultados.' },
  it: { appFindScanTruncated: 'La ricerca ha esaminato {cells} celle; questa cartella di lavoro grande potrebbe contenere altri risultati.' },
  pl: { appFindScanTruncated: 'Wyszukiwanie sprawdziło {cells} komórek; ten duży skoroszyt może zawierać więcej wyników.' },
  nl: { appFindScanTruncated: 'Zoeken heeft {cells} cellen gecontroleerd; deze grote werkmap kan meer resultaten bevatten.' },
  ms: { appFindScanTruncated: 'Carian mengimbas {cells} sel; buku kerja besar ini mungkin mempunyai hasil lain.' },
  he: { appFindScanTruncated: 'החיפוש סרק {cells} תאים; בחוברת עבודה גדולה זו עשויות להיות התאמות נוספות.' },
  hi: { appFindScanTruncated: 'खोज ने {cells} सेल जाँचे; इस बड़ी कार्यपुस्तिका में और परिणाम हो सकते हैं।' },
  'zh-TW': { appFindScanTruncated: '尋找已掃描 {cells} 個儲存格；活頁簿過大，結果可能不完整。' },
})
