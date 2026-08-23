import React, { useEffect, useMemo, useState } from 'react';
import {
  Alert,
  FlatList,
  Modal,
  SafeAreaView,
  ScrollView,
  StyleSheet,
  Text,
  TextInput,
  TouchableOpacity,
  View,
} from 'react-native';
import { CameraView, useCameraPermissions } from 'expo-camera';
import * as DocumentPicker from 'expo-document-picker';
import * as FileSystem from 'expo-file-system/legacy';
import * as SQLite from 'expo-sqlite';
import * as XLSX from 'xlsx';

type UnitType = 'ADET' | 'KOLI' | 'KUTU';

type Product = {
  barcode: string;
  code: string | null;
  name: string;
  unit: UnitType | null;
  unitsPerPack: number | null;
};

type ReceiptLine = Product & {
  enteredQty: number;
  baseQty: number;
};

const DB_NAME = 'kervan-mal-kabul.db';

function normalizeUnit(value: unknown): UnitType | null {
  const text = String(value ?? '').trim().toUpperCase();
  if (!text) return null;
  if (text.includes('KOL')) return 'KOLI';
  if (text.includes('KUT')) return 'KUTU';
  if (text.includes('ADET') || text === 'AD') return 'ADET';
  return null;
}

function parseNumber(value: unknown): number | null {
  const n = Number(String(value ?? '').replace(',', '.').trim());
  return Number.isFinite(n) && n > 0 ? n : null;
}

export default function App() {
  const [db, setDb] = useState<SQLite.SQLiteDatabase | null>(null);
  const [permission, requestPermission] = useCameraPermissions();
  const [screen, setScreen] = useState<'home' | 'scan' | 'history'>('home');
  const [scannedLocked, setScannedLocked] = useState(false);
  const [current, setCurrent] = useState<Product | null>(null);
  const [unknownBarcode, setUnknownBarcode] = useState<string | null>(null);
  const [qty, setQty] = useState('1');
  const [unit, setUnit] = useState<UnitType>('ADET');
  const [unitsPerPack, setUnitsPerPack] = useState('1');
  const [lines, setLines] = useState<ReceiptLine[]>([]);
  const [defModal, setDefModal] = useState(false);
  const [unknownName, setUnknownName] = useState('');
  const [unknownCode, setUnknownCode] = useState('');
  const [history, setHistory] = useState<any[]>([]);

  useEffect(() => {
    (async () => {
      const database = await SQLite.openDatabaseAsync(DB_NAME);
      await database.execAsync(`
        PRAGMA journal_mode = WAL;
        CREATE TABLE IF NOT EXISTS products (
          barcode TEXT PRIMARY KEY NOT NULL,
          code TEXT,
          name TEXT NOT NULL,
          unit TEXT,
          units_per_pack INTEGER,
          updated_at TEXT NOT NULL
        );
        CREATE TABLE IF NOT EXISTS receipts (
          id INTEGER PRIMARY KEY AUTOINCREMENT,
          created_at TEXT NOT NULL,
          total_lines INTEGER NOT NULL,
          total_base_qty INTEGER NOT NULL
        );
        CREATE TABLE IF NOT EXISTS receipt_lines (
          id INTEGER PRIMARY KEY AUTOINCREMENT,
          receipt_id INTEGER NOT NULL,
          barcode TEXT NOT NULL,
          code TEXT,
          name TEXT NOT NULL,
          unit TEXT NOT NULL,
          entered_qty REAL NOT NULL,
          units_per_pack REAL NOT NULL,
          base_qty REAL NOT NULL
        );
      `);
      setDb(database);
    })();
  }, []);

  const totalBaseQty = useMemo(
    () => lines.reduce((sum, x) => sum + x.baseQty, 0),
    [lines]
  );

  async function loadExcel() {
    if (!db) return;
    const picked = await DocumentPicker.getDocumentAsync({
      type: [
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'application/vnd.ms-excel',
        'text/csv',
      ],
      copyToCacheDirectory: true,
    });
    if (picked.canceled) return;

    try {
      const uri = picked.assets[0].uri;
      const base64 = await FileSystem.readAsStringAsync(uri, {
        encoding: FileSystem.EncodingType.Base64,
      });
      const workbook = XLSX.read(base64, { type: 'base64' });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json<Record<string, unknown>>(sheet, { defval: '' });

      let imported = 0;
      let skipped = 0;

      await db.withTransactionAsync(async () => {
        for (const row of rows) {
          const barcode = String(
            row['Barkod'] ?? row['BARKOD'] ?? row['Barcode'] ?? row['BARCODE'] ?? ''
          ).trim();
          const name = String(
            row['Ürün Adı'] ?? row['URUN ADI'] ?? row['Ürün'] ?? row['URUN'] ?? row['Stok Adı'] ?? row['STOK ADI'] ?? ''
          ).trim();
          const code = String(
            row['Ürün Kodu'] ?? row['URUN KODU'] ?? row['Stok Kodu'] ?? row['STOK KODU'] ?? ''
          ).trim() || null;
          const importedUnit = normalizeUnit(row['Birim'] ?? row['BIRIM'] ?? row['Birim Adı']);
          const importedPackQty = parseNumber(
            row['Koli İçi'] ?? row['KOLI ICI'] ?? row['Kutu İçi'] ?? row['KUTU ICI'] ?? row['Çevrim']
          );

          if (!barcode || !name) {
            skipped++;
            continue;
          }

          const existing = await db.getFirstAsync<Product>(
            'SELECT barcode, code, name, unit, units_per_pack as unitsPerPack FROM products WHERE barcode = ?',
            barcode
          );

          // Personelin uygulamada yaptığı birim/koli içi tanımını, Excel boş gelirse koru.
          const finalUnit = importedUnit ?? existing?.unit ?? null;
          const finalPackQty = importedPackQty ?? existing?.unitsPerPack ?? null;

          await db.runAsync(
            `INSERT INTO products(barcode, code, name, unit, units_per_pack, updated_at)
             VALUES (?, ?, ?, ?, ?, ?)
             ON CONFLICT(barcode) DO UPDATE SET
               code = excluded.code,
               name = excluded.name,
               unit = excluded.unit,
               units_per_pack = excluded.units_per_pack,
               updated_at = excluded.updated_at`,
            barcode,
            code,
            name,
            finalUnit,
            finalPackQty,
            new Date().toISOString()
          );
          imported++;
        }
      });

      Alert.alert('Ürün listesi güncellendi', `${imported} ürün alındı. ${skipped} satır atlandı.`);
    } catch (e) {
      console.error(e);
      Alert.alert('Excel okunamadı', 'Dosya sütunlarını kontrol edin. Barkod ve ürün adı zorunludur.');
    }
  }

  async function findProduct(barcode: string) {
    if (!db) return;
    const product = await db.getFirstAsync<Product>(
      `SELECT barcode, code, name, unit, units_per_pack as unitsPerPack
       FROM products WHERE barcode = ?`,
      barcode
    );

    if (!product) {
      setUnknownBarcode(barcode);
      setCurrent(null);
      setUnknownName('');
      setUnknownCode('');
      setUnit('ADET');
      setUnitsPerPack('1');
      setDefModal(true);
      return;
    }

    setCurrent(product);
    setQty('1');
    setUnit(product.unit ?? 'ADET');
    setUnitsPerPack(String(product.unitsPerPack ?? 1));

    if (!product.unit || !product.unitsPerPack) {
      setDefModal(true);
    }
  }

  async function saveDefinition() {
    if (!db) return;
    const barcode = current?.barcode ?? unknownBarcode;
    if (!barcode) return;

    const packQty = unit === 'ADET' ? 1 : Number(unitsPerPack);
    if (!Number.isFinite(packQty) || packQty <= 0) {
      Alert.alert('Hatalı miktar', 'Koli/kutu içi adet 1 veya daha büyük olmalı.');
      return;
    }

    if (!current && !unknownName.trim()) {
      Alert.alert('Ürün adı gerekli', 'Tanımsız barkod için ürün adını girin.');
      return;
    }

    const product: Product = {
      barcode,
      code: current?.code ?? (unknownCode.trim() || null),
      name: current?.name ?? unknownName.trim(),
      unit,
      unitsPerPack: packQty,
    };

    await db.runAsync(
      `INSERT INTO products(barcode, code, name, unit, units_per_pack, updated_at)
       VALUES (?, ?, ?, ?, ?, ?)
       ON CONFLICT(barcode) DO UPDATE SET
         code = excluded.code,
         name = excluded.name,
         unit = excluded.unit,
         units_per_pack = excluded.units_per_pack,
         updated_at = excluded.updated_at`,
      product.barcode,
      product.code,
      product.name,
      product.unit,
      product.unitsPerPack,
      new Date().toISOString()
    );

    setCurrent(product);
    setUnknownBarcode(null);
    setDefModal(false);
  }

  async function updateCurrentDefinition() {
    if (!db || !current) return;
    const packQty = unit === 'ADET' ? 1 : Number(unitsPerPack);
    if (!Number.isFinite(packQty) || packQty <= 0) {
      Alert.alert('Hatalı miktar', 'Koli/kutu içi adet geçerli değil.');
      return;
    }
    await db.runAsync(
      'UPDATE products SET unit = ?, units_per_pack = ?, updated_at = ? WHERE barcode = ?',
      unit,
      packQty,
      new Date().toISOString(),
      current.barcode
    );
    setCurrent({ ...current, unit, unitsPerPack: packQty });
  }

  async function addLine() {
    if (!current) return;
    const entered = Number(qty);
    const pack = unit === 'ADET' ? 1 : Number(unitsPerPack);
    if (!Number.isFinite(entered) || entered <= 0 || !Number.isFinite(pack) || pack <= 0) {
      Alert.alert('Miktarı kontrol edin');
      return;
    }

    await updateCurrentDefinition();
    const baseQty = entered * pack;

    setLines(prev => {
      const idx = prev.findIndex(x => x.barcode === current.barcode && x.unit === unit && x.unitsPerPack === pack);
      if (idx >= 0) {
        const copy = [...prev];
        copy[idx] = {
          ...copy[idx],
          enteredQty: copy[idx].enteredQty + entered,
          baseQty: copy[idx].baseQty + baseQty,
        };
        return copy;
      }
      return [
        ...prev,
        {
          ...current,
          unit,
          unitsPerPack: pack,
          enteredQty: entered,
          baseQty,
        },
      ];
    });

    setCurrent(null);
    setQty('1');
    setScannedLocked(false);
  }

  async function finishReceipt() {
    if (!db || lines.length === 0) return;
    const result = await db.runAsync(
      'INSERT INTO receipts(created_at, total_lines, total_base_qty) VALUES (?, ?, ?)',
      new Date().toISOString(),
      lines.length,
      totalBaseQty
    );
    const receiptId = result.lastInsertRowId;
    for (const line of lines) {
      await db.runAsync(
        `INSERT INTO receipt_lines(
          receipt_id, barcode, code, name, unit, entered_qty, units_per_pack, base_qty
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)`,
        receiptId,
        line.barcode,
        line.code,
        line.name,
        line.unit ?? 'ADET',
        line.enteredQty,
        line.unitsPerPack ?? 1,
        line.baseQty
      );
    }
    Alert.alert('Mal kabul tamamlandı', `${lines.length} kalem / ${totalBaseQty} adet karşılığı kaydedildi.`);
    setLines([]);
    setCurrent(null);
    setScreen('home');
  }

  async function loadHistory() {
    if (!db) return;
    const rows = await db.getAllAsync(
      'SELECT id, created_at, total_lines, total_base_qty FROM receipts ORDER BY id DESC LIMIT 100'
    );
    setHistory(rows);
    setScreen('history');
  }

  if (!db) {
    return <SafeAreaView style={styles.center}><Text>Veritabanı hazırlanıyor…</Text></SafeAreaView>;
  }

  return (
    <SafeAreaView style={styles.safe}>
      {screen === 'home' && (
        <View style={styles.page}>
          <Text style={styles.title}>Kervan Market</Text>
          <Text style={styles.subtitle}>Mal Kabul</Text>

          <TouchableOpacity style={styles.primary} onPress={() => setScreen('scan')}>
            <Text style={styles.primaryText}>📦 Yeni Mal Kabul</Text>
          </TouchableOpacity>
          <TouchableOpacity style={styles.secondary} onPress={loadExcel}>
            <Text style={styles.secondaryText}>📄 AKINSOFT Excel / CSV Yükle</Text>
          </TouchableOpacity>
          <TouchableOpacity style={styles.secondary} onPress={loadHistory}>
            <Text style={styles.secondaryText}>📋 Mal Kabul Geçmişi</Text>
          </TouchableOpacity>

          <View style={styles.infoBox}>
            <Text style={styles.infoTitle}>Birim sistemi</Text>
            <Text style={styles.info}>ADET → 1 = 1 adet</Text>
            <Text style={styles.info}>KOLİ → koli sayısı × koli içi adet</Text>
            <Text style={styles.info}>KUTU → kutu sayısı × kutu içi adet</Text>
          </View>
        </View>
      )}

      {screen === 'scan' && (
        <View style={styles.page}>
          <View style={styles.rowBetween}>
            <TouchableOpacity onPress={() => setScreen('home')}><Text style={styles.link}>← Ana Sayfa</Text></TouchableOpacity>
            <Text style={styles.counter}>{lines.length} kalem / {totalBaseQty} adet</Text>
          </View>

          {!current && (
            <View style={styles.cameraWrap}>
              {!permission?.granted ? (
                <TouchableOpacity style={styles.primary} onPress={requestPermission}>
                  <Text style={styles.primaryText}>Kamera İzni Ver</Text>
                </TouchableOpacity>
              ) : (
                <CameraView
                  style={styles.camera}
                  facing="back"
                  barcodeScannerSettings={{
                    barcodeTypes: ['ean13', 'ean8', 'upc_a', 'upc_e', 'code128', 'code39'],
                  }}
                  onBarcodeScanned={scannedLocked ? undefined : async ({ data }) => {
                    setScannedLocked(true);
                    await findProduct(String(data).trim());
                  }}
                />
              )}
              <Text style={styles.scanHint}>Barkodu kameraya gösterin</Text>
              {scannedLocked && !current && !defModal && (
                <TouchableOpacity onPress={() => setScannedLocked(false)}><Text style={styles.link}>Tekrar tara</Text></TouchableOpacity>
              )}
            </View>
          )}

          {current && (
            <ScrollView contentContainerStyle={styles.card}>
              <Text style={styles.productName}>{current.name}</Text>
              <Text style={styles.muted}>Barkod: {current.barcode}</Text>
              {!!current.code && <Text style={styles.muted}>Kod: {current.code}</Text>}

              <Text style={styles.label}>Geliş birimi</Text>
              <View style={styles.unitRow}>
                {(['ADET', 'KOLI', 'KUTU'] as UnitType[]).map(x => (
                  <TouchableOpacity key={x} style={[styles.unitBtn, unit === x && styles.unitBtnActive]} onPress={() => {
                    setUnit(x);
                    if (x === 'ADET') setUnitsPerPack('1');
                  }}>
                    <Text style={[styles.unitText, unit === x && styles.unitTextActive]}>{x}</Text>
                  </TouchableOpacity>
                ))}
              </View>

              {unit !== 'ADET' && (
                <>
                  <Text style={styles.label}>{unit === 'KOLI' ? '1 koli kaç adet?' : '1 kutu kaç adet?'}</Text>
                  <TextInput
                    value={unitsPerPack}
                    onChangeText={setUnitsPerPack}
                    keyboardType="numeric"
                    style={styles.input}
                    placeholder="Örn. 24"
                  />
                </>
              )}

              <Text style={styles.label}>Gelen {unit.toLocaleLowerCase('tr-TR')} sayısı</Text>
              <TextInput value={qty} onChangeText={setQty} keyboardType="decimal-pad" style={styles.inputBig} />

              <View style={styles.resultBox}>
                <Text style={styles.resultText}>
                  {qty || '0'} {unit.toLocaleLowerCase('tr-TR')} = {Number(qty || 0) * Number(unitsPerPack || 1)} adet
                </Text>
              </View>

              <TouchableOpacity style={styles.primary} onPress={addLine}>
                <Text style={styles.primaryText}>✓ Mal Kabule Ekle</Text>
              </TouchableOpacity>
              <TouchableOpacity onPress={() => { setCurrent(null); setScannedLocked(false); }}>
                <Text style={styles.cancel}>İptal / tekrar tara</Text>
              </TouchableOpacity>
            </ScrollView>
          )}

          {lines.length > 0 && !current && (
            <View style={styles.bottomBox}>
              <FlatList
                style={{ maxHeight: 160 }}
                data={[...lines].reverse()}
                keyExtractor={(x, i) => `${x.barcode}-${i}`}
                renderItem={({ item }) => (
                  <View style={styles.lineRow}>
                    <View style={{ flex: 1 }}>
                      <Text numberOfLines={1} style={styles.lineName}>{item.name}</Text>
                      <Text style={styles.muted}>{item.enteredQty} {item.unit?.toLocaleLowerCase('tr-TR')} = {item.baseQty} adet</Text>
                    </View>
                  </View>
                )}
              />
              <TouchableOpacity style={styles.finish} onPress={finishReceipt}>
                <Text style={styles.primaryText}>Mal Kabulü Tamamla</Text>
              </TouchableOpacity>
            </View>
          )}
        </View>
      )}

      {screen === 'history' && (
        <View style={styles.page}>
          <TouchableOpacity onPress={() => setScreen('home')}><Text style={styles.link}>← Ana Sayfa</Text></TouchableOpacity>
          <Text style={styles.titleSmall}>Mal Kabul Geçmişi</Text>
          <FlatList
            data={history}
            keyExtractor={(x) => String(x.id)}
            ListEmptyComponent={<Text style={styles.muted}>Henüz kayıt yok.</Text>}
            renderItem={({ item }) => (
              <View style={styles.historyCard}>
                <Text style={styles.lineName}>Mal Kabul #{item.id}</Text>
                <Text>{new Date(item.created_at).toLocaleString('tr-TR')}</Text>
                <Text>{item.total_lines} kalem • {item.total_base_qty} adet karşılığı</Text>
              </View>
            )}
          />
        </View>
      )}

      <Modal transparent visible={defModal} animationType="slide" onRequestClose={() => setDefModal(false)}>
        <View style={styles.modalBackdrop}>
          <ScrollView contentContainerStyle={styles.modalCard}>
            <Text style={styles.modalTitle}>{current ? 'Birim Tanımı Eksik' : 'Tanımsız Barkod'}</Text>
            <Text style={styles.muted}>Barkod: {current?.barcode ?? unknownBarcode}</Text>

            {!current && (
              <>
                <Text style={styles.warning}>Bu barkod AKINSOFT listesinden bulunamadı. Personel ürün kartını tanımlayabilir.</Text>
                <Text style={styles.label}>Ürün adı</Text>
                <TextInput value={unknownName} onChangeText={setUnknownName} style={styles.input} placeholder="Ürün adı" />
                <Text style={styles.label}>Ürün kodu (opsiyonel)</Text>
                <TextInput value={unknownCode} onChangeText={setUnknownCode} style={styles.input} placeholder="Stok kodu" />
              </>
            )}

            <Text style={styles.label}>Birim</Text>
            <View style={styles.unitRow}>
              {(['ADET', 'KOLI', 'KUTU'] as UnitType[]).map(x => (
                <TouchableOpacity key={x} style={[styles.unitBtn, unit === x && styles.unitBtnActive]} onPress={() => {
                  setUnit(x);
                  if (x === 'ADET') setUnitsPerPack('1');
                }}>
                  <Text style={[styles.unitText, unit === x && styles.unitTextActive]}>{x}</Text>
                </TouchableOpacity>
              ))}
            </View>

            {unit !== 'ADET' && (
              <>
                <Text style={styles.label}>1 {unit.toLocaleLowerCase('tr-TR')} kaç adet?</Text>
                <TextInput value={unitsPerPack} onChangeText={setUnitsPerPack} keyboardType="numeric" style={styles.inputBig} />
              </>
            )}

            <TouchableOpacity style={styles.primary} onPress={saveDefinition}>
              <Text style={styles.primaryText}>Kaydet ve Devam Et</Text>
            </TouchableOpacity>
            <TouchableOpacity onPress={() => {
              setDefModal(false);
              setUnknownBarcode(null);
              setCurrent(null);
              setScannedLocked(false);
            }}>
              <Text style={styles.cancel}>Vazgeç</Text>
            </TouchableOpacity>
          </ScrollView>
        </View>
      </Modal>
    </SafeAreaView>
  );
}

const styles = StyleSheet.create({
  safe: { flex: 1, backgroundColor: '#f6f7f9' },
  page: { flex: 1, padding: 18, gap: 14 },
  center: { flex: 1, alignItems: 'center', justifyContent: 'center' },
  title: { fontSize: 32, fontWeight: '900', marginTop: 18 },
  subtitle: { fontSize: 20, fontWeight: '700', marginBottom: 20 },
  titleSmall: { fontSize: 25, fontWeight: '900', marginVertical: 15 },
  primary: { backgroundColor: '#ef7d00', padding: 17, borderRadius: 14, alignItems: 'center' },
  primaryText: { color: '#fff', fontWeight: '900', fontSize: 17 },
  secondary: { backgroundColor: '#0b2a4a', padding: 17, borderRadius: 14, alignItems: 'center' },
  secondaryText: { color: '#fff', fontWeight: '800', fontSize: 15 },
  infoBox: { backgroundColor: '#fff', padding: 16, borderRadius: 14, marginTop: 12 },
  infoTitle: { fontWeight: '900', fontSize: 16, marginBottom: 8 },
  info: { marginVertical: 3 },
  rowBetween: { flexDirection: 'row', justifyContent: 'space-between', alignItems: 'center' },
  link: { color: '#0b63ce', fontWeight: '800' },
  counter: { fontWeight: '800' },
  cameraWrap: { flex: 1, gap: 10 },
  camera: { flex: 1, minHeight: 330, borderRadius: 18, overflow: 'hidden' },
  scanHint: { textAlign: 'center', fontSize: 16, fontWeight: '700' },
  card: { backgroundColor: '#fff', borderRadius: 18, padding: 18, gap: 12 },
  productName: { fontSize: 24, fontWeight: '900' },
  muted: { color: '#68707a' },
  label: { fontWeight: '800', marginTop: 4 },
  unitRow: { flexDirection: 'row', gap: 8 },
  unitBtn: { flex: 1, borderWidth: 1, borderColor: '#cdd3da', borderRadius: 12, paddingVertical: 13, alignItems: 'center' },
  unitBtnActive: { backgroundColor: '#0b2a4a', borderColor: '#0b2a4a' },
  unitText: { fontWeight: '800' },
  unitTextActive: { color: '#fff' },
  input: { borderWidth: 1, borderColor: '#cfd5dc', borderRadius: 12, padding: 13, backgroundColor: '#fff' },
  inputBig: { borderWidth: 1, borderColor: '#cfd5dc', borderRadius: 12, padding: 14, backgroundColor: '#fff', fontSize: 25, fontWeight: '900', textAlign: 'center' },
  resultBox: { backgroundColor: '#eef6ff', borderRadius: 12, padding: 14 },
  resultText: { textAlign: 'center', fontWeight: '900', fontSize: 18 },
  cancel: { textAlign: 'center', color: '#ad1d1d', fontWeight: '800', padding: 10 },
  bottomBox: { backgroundColor: '#fff', borderRadius: 16, padding: 12, gap: 10 },
  lineRow: { paddingVertical: 8, borderBottomWidth: StyleSheet.hairlineWidth, borderBottomColor: '#ddd' },
  lineName: { fontWeight: '800' },
  finish: { backgroundColor: '#176b32', padding: 15, borderRadius: 13, alignItems: 'center' },
  historyCard: { backgroundColor: '#fff', borderRadius: 13, padding: 15, marginBottom: 10, gap: 4 },
  modalBackdrop: { flex: 1, backgroundColor: 'rgba(0,0,0,.45)', justifyContent: 'flex-end' },
  modalCard: { backgroundColor: '#fff', borderTopLeftRadius: 22, borderTopRightRadius: 22, padding: 20, gap: 12 },
  modalTitle: { fontSize: 23, fontWeight: '900' },
  warning: { backgroundColor: '#fff3cd', padding: 12, borderRadius: 10, fontWeight: '700' },
});
