"use client";

import React, { useEffect, useMemo, useState } from "react";
import {
  Activity,
  AlertCircle,
  ArrowRightLeft,
  BarChart3,
  CalendarDays,
  CheckCircle2,
  ClipboardList,
  DollarSign,
  FileSpreadsheet,
  Loader2,
  MapPin,
  RefreshCw,
  ShieldAlert,
  TrendingUp,
  Upload,
  Users,
  Zap,
} from "lucide-react";
import * as XLSX from "xlsx";

const API_BASE = process.env.NEXT_PUBLIC_API_URL || "http://127.0.0.1:8000";
const COSTO_POR_HORA_PROMEDIO = 46.5;
const AXH_META = 7.0;

const DIAS = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"];
const AREAS = ["Taquilla", "Dulcería", "Lobby", "Baños", "Entrada"];

const DISTRIBUCION_DEMANDA: { [key: string]: { [hora: string]: number } } = {
  Miércoles: { "09:00": 0, "10:00": 0.0079, "11:00": 0.0109, "12:00": 0.0311, "13:00": 0.0305, "14:00": 0.0456, "15:00": 0.0817, "16:00": 0.0781, "17:00": 0.0987, "18:00": 0.1519, "19:00": 0.1788, "20:00": 0.1238, "21:00": 0.1155, "22:00": 0.0454, "23:00": 0 },
  Jueves: { "09:00": 0, "10:00": 0.0008, "11:00": 0.0114, "12:00": 0.0373, "13:00": 0.0346, "14:00": 0.0517, "15:00": 0.0868, "16:00": 0.0724, "17:00": 0.1081, "18:00": 0.1518, "19:00": 0.152, "20:00": 0.1329, "21:00": 0.1152, "22:00": 0.045, "23:00": 0 },
  Viernes: { "09:00": 0, "10:00": 0.0006, "11:00": 0.0085, "12:00": 0.0308, "13:00": 0.0302, "14:00": 0.0424, "15:00": 0.0734, "16:00": 0.0735, "17:00": 0.1007, "18:00": 0.1412, "19:00": 0.1455, "20:00": 0.1492, "21:00": 0.147, "22:00": 0.0571, "23:00": 0 },
  Sábado: { "09:00": 0.0014, "10:00": 0.0095, "11:00": 0.0148, "12:00": 0.0289, "13:00": 0.038, "14:00": 0.0519, "15:00": 0.0974, "16:00": 0.0896, "17:00": 0.1354, "18:00": 0.1613, "19:00": 0.1138, "20:00": 0.1172, "21:00": 0.1029, "22:00": 0.0379, "23:00": 0 },
  Domingo: { "09:00": 0.0023, "10:00": 0.0175, "11:00": 0.0225, "12:00": 0.0499, "13:00": 0.0611, "14:00": 0.0755, "15:00": 0.1201, "16:00": 0.1003, "17:00": 0.135, "18:00": 0.1531, "19:00": 0.0996, "20:00": 0.0857, "21:00": 0.0594, "22:00": 0.0179, "23:00": 0 },
  Lunes: { "09:00": 0, "10:00": 0.004, "11:00": 0.0095, "12:00": 0.0302, "13:00": 0.0307, "14:00": 0.049, "15:00": 0.0777, "16:00": 0.0727, "17:00": 0.1036, "18:00": 0.1387, "19:00": 0.1655, "20:00": 0.1456, "21:00": 0.1222, "22:00": 0.0505, "23:00": 0 },
  Martes: { "09:00": 0, "10:00": 0.0014, "11:00": 0.0107, "12:00": 0.0384, "13:00": 0.0311, "14:00": 0.0542, "15:00": 0.0735, "16:00": 0.077, "17:00": 0.104, "18:00": 0.1456, "19:00": 0.1787, "20:00": 0.1309, "21:00": 0.1137, "22:00": 0.0408, "23:00": 0 },
};

type Vista = "mapa" | "plan-semanal" | "analisis" | "asistencia";

type EmpleadoLocal = {
  id_ps: string;
  id?: string | number;
  nombre: string;
  area: string;
  nivel_esp: number;
  expertise: Record<string, number>;
  turnos: Record<string, { e: string; s: string; descanso: boolean }>;
};

const fixText = (value: any) => {
  let s = String(value ?? "");
  const rep: Record<string, string> = {
    "DÃ­a": "Día",
    "invÃ¡lido": "inválido",
    "MiÃ©rcoles": "Miércoles",
    "SÃ¡bado": "Sábado",
    "DulcerÃ­a": "Dulcería",
    "BaÃ±os": "Baños",
    "OptimizaciÃ³n": "Optimización",
    "Ã¡": "á",
    "Ã©": "é",
    "Ã­": "í",
    "Ã³": "ó",
    "Ãº": "ú",
    "Ã±": "ñ",
  };
  Object.entries(rep).forEach(([bad, good]) => {
    s = s.replaceAll(bad, good);
  });
  return s;
};

const idOf = (cp: any) => String(cp?.id_ps ?? cp?.id ?? "");

const money = (n: number) =>
  Number(n || 0).toLocaleString("es-MX", { style: "currency", currency: "MXN" });

const estadoColor = (estado: string) => {
  const e = fixText(estado).toLowerCase();
  if (e.includes("falta")) return "bg-rose-100 text-rose-700 border-rose-200";
  if (e.includes("ley silla")) return "bg-amber-100 text-amber-700 border-amber-200";
  if (e.includes("alimentos")) return "bg-orange-100 text-orange-700 border-orange-200";
  if (e.includes("descanso")) return "bg-sky-100 text-sky-700 border-sky-200";
  if (e.includes("reasignado")) return "bg-violet-100 text-violet-700 border-violet-200";
  if (e.includes("operando")) return "bg-emerald-100 text-emerald-700 border-emerald-200";
  return "bg-slate-100 text-slate-600 border-slate-200";
};

const areaCellColor = (area: string) => {
  const a = fixText(area);
  if (a === "Taquilla") return "bg-blue-50 text-blue-900 border-blue-100";
  if (a === "Dulcería") return "bg-emerald-50 text-emerald-900 border-emerald-100";
  if (a === "Lobby") return "bg-violet-50 text-violet-900 border-violet-100";
  if (a === "Baños") return "bg-amber-50 text-amber-900 border-amber-100";
  if (a === "Entrada") return "bg-cyan-50 text-cyan-900 border-cyan-100";
  return "bg-slate-50 text-slate-500 border-slate-100";
};

const barWidth = (value: number, max: number) => `${Math.min(100, (Number(value || 0) / Math.max(1, max)) * 100)}%`;

export default function CineOpsDashboard() {
  const [subVista, setSubVista] = useState<Vista>("mapa");
  const [horaSeleccionada, setHoraSeleccionada] = useState("16:00");
  const [diaSeleccionada, setDiaSeleccionada] = useState("Sábado");

  const [demandaSemanaActual, setDemandaSemanaActual] = useState(7200);
  const [presupuestoHorasMax, setPresupuestoHorasMax] = useState(3032);

  const [faltasHoy, setFaltasHoy] = useState<Record<string, boolean>>({});
  const [descansosHoy, setDescansosHoy] = useState<Record<string, boolean>>({});
  const [nombreArchivo, setNombreArchivo] = useState<string | null>(null);
  const [estadoCarga, setEstadoCarga] = useState<"idle" | "leyendo" | "exito" | "error">("idle");
  const [mensajeDSS, setMensajeDSS] = useState("");
  const [kpisBackend, setKpisBackend] = useState<any>(null);
  const [optimizando, setOptimizando] = useState(false);
  const [descansoAutoDias, setDescansoAutoDias] = useState<1 | 2>(1);

  const [cinepolitosMatriz, setCinepolitosMatriz] = useState<EmpleadoLocal[]>(() => {
    const areasMock = ["Taquilla", "Dulcería", "Lobby", "Baños", "Entrada"];
    const lista: EmpleadoLocal[] = [];
    for (let i = 1; i <= 78; i++) {
      const areaAsignada = areasMock[i % 5];
      const lvl = i > 41 ? 3 : i > 18 ? 2 : 1;
      lista.push({
        id_ps: String(204100 + i),
        nombre: `CINEPOLITO GENERAL EXTRA ${i}`,
        area: areaAsignada,
        nivel_esp: lvl,
        expertise: {
          M1: areaAsignada === "Taquilla" || i % 3 === 0 ? 1 : 0,
          M2: areaAsignada === "Dulcería" || i % 4 === 0 ? 1 : 0,
          M3: areaAsignada === "Lobby" || i % 2 === 0 ? 1 : 0,
          M4: areaAsignada === "Baños" || i % 5 === 0 ? 1 : 0,
          M5: areaAsignada === "Entrada" || i % 6 === 0 ? 1 : 0,
        },
        turnos: {
          Lunes: { e: "15:00", s: "22:00", descanso: false },
          Martes: { e: "15:00", s: "22:00", descanso: false },
          Miércoles: { e: "15:00", s: "22:00", descanso: false },
          Jueves: { e: "15:00", s: "22:00", descanso: false },
          Viernes: { e: "15:00", s: "22:00", descanso: false },
          Sábado: { e: i % 7 === 0 ? "Descanso" : "14:00", s: i % 7 === 0 ? "Descanso" : "22:00", descanso: i % 7 === 0 },
          Domingo: { e: "12:00", s: "20:00", descanso: false },
        },
      });
    }
    return lista;
  });

  const asistenciasParaAPI = () => {
    const asistencias: Record<string, string> = {};
    Object.entries(faltasHoy).forEach(([id, marcado]) => {
      if (marcado) asistencias[id] = "FALTA";
    });
    return asistencias;
  };

  const refrescarKPIsBackend = async () => {
    try {
      const res = await fetch(`${API_BASE}/kpis`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          dia: diaSeleccionada,
          hora: horaSeleccionada,
          asistencias: asistenciasParaAPI(),
        }),
      });
      const data = await res.json().catch(() => ({}));
      if (!res.ok) {
        setMensajeDSS(fixText(data.detail || "No se pudieron actualizar los KPIs."));
        return;
      }
      setKpisBackend(data);
      if (data.ppto_horas_sem) setPresupuestoHorasMax(Number(data.ppto_horas_sem));
      setMensajeDSS((m) => (m.includes("Error") || m.includes("inválido") ? "" : m));
    } catch {
      // Fallback local cuando el backend no está corriendo.
    }
  };

  useEffect(() => {
    refrescarKPIsBackend();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [diaSeleccionada, horaSeleccionada, faltasHoy, descansosHoy]);

  const subirExcelAlBackend = async (file: File) => {
    const formData = new FormData();
    formData.append("file", file);
    const res = await fetch(`${API_BASE}/upload`, { method: "POST", body: formData });
    const data = await res.json().catch(() => ({}));
    if (!res.ok) throw new Error(fixText(data.detail || "No se pudo procesar el archivo en backend."));
    await refrescarKPIsBackend();
  };

  const recalcularOptimizacion = async (descansosAutomaticos = false) => {
    setOptimizando(true);
    setMensajeDSS(descansosAutomaticos ? "Recalculando solución con descansos automáticos..." : "Recalculando optimización...");
    try {
      const res = await fetch(`${API_BASE}/optimize-weekly`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          tipo_semana: "Normal",
          hora_seleccionada: horaSeleccionada,
          descansos_automaticos: descansosAutomaticos,
          descanso_dias_objetivo: descansosAutomaticos ? descansoAutoDias : 0,
        }),
      });
      const data = await res.json().catch(() => ({}));
      if (!res.ok) throw new Error(fixText(data.detail || "No se pudo optimizar."));
      setMensajeDSS(`Optimización recalculada correctamente. Estado: ${data.solver_status}. Horas: ${data.horas_asignadas_sem} / ${data.presupuesto_horas}.`);
      await refrescarKPIsBackend();
    } catch (err: any) {
      setMensajeDSS(fixText(err.message || "Error al optimizar."));
    } finally {
      setOptimizando(false);
    }
  };

  const marcarFaltaBackend = async (cp: any) => {
    const id = idOf(cp);
    const nuevoValor = !faltasHoy[id];
    setFaltasHoy((p) => ({ ...p, [id]: nuevoValor }));
    try {
      const res = await fetch(`${API_BASE}/attendance`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ empleado_id: id, dia: diaSeleccionada, estado: nuevoValor ? "FALTA" : "PRESENTE" }),
      });
      const data = await res.json().catch(() => ({}));
      if (!res.ok) throw new Error(fixText(data.detail || "No se pudo actualizar la falta."));
      setMensajeDSS(nuevoValor ? `Falta registrada: ${fixText(cp.nombre)}. El solver liberó sus horas.` : `Falta retirada: ${fixText(cp.nombre)}.`);
      await refrescarKPIsBackend();
    } catch (err: any) {
      setMensajeDSS(fixText(err.message || "Error al registrar asistencia."));
    }
  };

  const marcarDescansoManualBackend = async (cp: any) => {
    const id = idOf(cp);
    const limpiar = !!descansosHoy[id];
    try {
      const res = await fetch(`${API_BASE}/rest-days`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ empleado_id: id, dias: limpiar ? [] : [diaSeleccionada], modo: limpiar ? "limpiar" : "manual" }),
      });
      const data = await res.json().catch(() => ({}));
      if (!res.ok) throw new Error(fixText(data.detail || "No se pudo marcar descanso."));
      setDescansosHoy((p) => ({ ...p, [id]: !limpiar }));
      setMensajeDSS(limpiar ? `Descanso retirado: ${fixText(cp.nombre)}.` : `Descanso manual aplicado a ${fixText(cp.nombre)}. KPIs actualizados.`);
      await refrescarKPIsBackend();
    } catch (err: any) {
      setMensajeDSS(fixText(err.message || "Error al marcar descanso."));
    }
  };

  const aplicarRecomendacion = async (r: any) => {
    try {
      const res = await fetch(`${API_BASE}/manual-assignment`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ empleado_id: r.id, dia: diaSeleccionada, hora: horaSeleccionada, area: r.mover_a }),
      });
      const data = await res.json().catch(() => ({}));
      if (!res.ok) throw new Error(fixText(data.detail || "No se pudo aplicar la recomendación."));
      setMensajeDSS(`Recomendación aplicada: ${fixText(r.empleado)} → ${fixText(r.mover_a)}.`);
      await refrescarKPIsBackend();
    } catch (err: any) {
      setMensajeDSS(fixText(err.message || "Error al aplicar recomendación."));
    }
  };

  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setNombreArchivo(file.name);
    setEstadoCarga("leyendo");
    subirExcelAlBackend(file).catch((err) => setMensajeDSS(fixText(err.message || "El backend no respondió; se mantiene lectura local.")));

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const workbook = XLSX.read(evt.target?.result, { type: "binary" });
        const nombreHojaDemanda = workbook.SheetNames.find((name) => name.toLowerCase().includes("demanda"));
        if (nombreHojaDemanda) {
          const rowsDemanda: any[] = XLSX.utils.sheet_to_json(workbook.Sheets[nombreHojaDemanda], { header: 1 });
          rowsDemanda.forEach((row) => {
            const txt = JSON.stringify(row).toLowerCase();
            const num = row.find((cell: any) => typeof cell === "number");
            if ((txt.includes("demanda") || txt.includes("semana actual") || txt.includes("admit")) && num) setDemandaSemanaActual(num);
            if ((txt.includes("presupuesto") || txt.includes("horas limite") || txt.includes("horas límite")) && num) setPresupuestoHorasMax(num);
          });
        }

        const nombreHojaPlantilla = workbook.SheetNames.find((name) => !name.toLowerCase().includes("demanda")) || workbook.SheetNames[0];
        const rowsRaw: any[][] = XLSX.utils.sheet_to_json(workbook.Sheets[nombreHojaPlantilla], { header: 1 });
        let indiceEncabezado = -1;
        for (let i = 0; i < rowsRaw.length; i++) {
          const filaUnida = rowsRaw[i].join(" ").toLowerCase();
          if (filaUnida.includes("id") && (filaUnida.includes("nombre") || filaUnida.includes("colaborador") || filaUnida.includes("turno"))) {
            indiceEncabezado = i;
            break;
          }
        }
        if (indiceEncabezado === -1) {
          setEstadoCarga("error");
          return;
        }

        const headers = rowsRaw[indiceEncabezado].map((h) => String(h || "").toLowerCase().trim());
        const findIndexByVariants = (variants: string[]) => headers.findIndex((h) => variants.some((v) => h.includes(v)));
        const idxId = findIndexByVariants(["id", "ps", "expediente"]);
        const idxNombre = findIndexByVariants(["nombre", "colaborador", "cinepolito"]);
        const idxNivel = findIndexByVariants(["nivel", "expertis"]);
        const idxM1 = findIndexByVariants(["m1"]);
        const idxM2 = findIndexByVariants(["m2"]);
        const idxM3 = findIndexByVariants(["m3"]);
        const idxM4 = findIndexByVariants(["m4"]);
        const idxM5 = findIndexByVariants(["m5"]);
        const idxDias = DIAS.map((d) => headers.findIndex((h) => h.includes(d.toLowerCase())));

        const analizados: EmpleadoLocal[] = [];
        for (let i = indiceEncabezado + 1; i < rowsRaw.length; i++) {
          const fila = rowsRaw[i];
          if (!fila || !fila[idxNombre] || String(fila[idxNombre]).trim() === "") continue;
          const id_ps = fila[idxId] ? String(fila[idxId]).trim() : `ID-${200000 + i}`;
          const nombre = fixText(String(fila[idxNombre])).toUpperCase().trim();
          if (nombre.includes("TOTAL") || nombre.includes("CONTROL DE")) continue;
          const nivelEsp = parseInt(fila[idxNivel] || "1", 10) || 1;
          const m1 = parseInt(fila[idxM1] || "0", 10) === 1 ? 1 : 0;
          const m2 = parseInt(fila[idxM2] || "0", 10) === 1 ? 1 : 0;
          const m3 = parseInt(fila[idxM3] || "0", 10) === 1 ? 1 : 0;
          const m4 = parseInt(fila[idxM4] || "0", 10) === 1 ? 1 : 0;
          const m5 = parseInt(fila[idxM5] || "0", 10) === 1 ? 1 : 0;

          let area = "Lobby";
          if (m1) area = "Taquilla";
          else if (m2) area = "Dulcería";
          else if (m3) area = "Lobby";
          else if (m4) area = "Baños";
          else if (m5) area = "Entrada";

          const turnos: any = {};
          DIAS.forEach((d, dIdx) => {
            const cellRaw = fila[idxDias[dIdx]];
            if (!cellRaw || /descanso|x/i.test(String(cellRaw)) || String(cellRaw).trim() === "0") {
              turnos[d] = { e: "Descanso", s: "Descanso", descanso: true };
            } else {
              const arr = String(cellRaw).split("-");
              turnos[d] = arr.length === 2 ? { e: arr[0].trim(), s: arr[1].trim(), descanso: false } : { e: "15:00", s: "22:00", descanso: false };
            }
          });

          analizados.push({ id_ps, nombre, area, nivel_esp: nivelEsp, expertise: { M1: m1, M2: m2, M3: m3, M4: m4, M5: m5 }, turnos });
        }
        if (analizados.length > 0) {
          setCinepolitosMatriz(analizados);
          setEstadoCarga("exito");
        } else setEstadoCarga("error");
      } catch {
        setEstadoCarga("error");
      }
    };
    reader.readAsBinaryString(file);
  };

  const calcularHorasIndividuales = (turnos: any) => {
    let suma = 0;
    DIAS.forEach((d) => {
      const t = turnos?.[d];
      if (t && !t.descanso && t.e !== "Descanso") {
        const [he, me] = t.e.split(":").map(Number);
        const [hs, ms] = t.s.split(":").map(Number);
        let diff = hs + ms / 60 - (he + me / 60);
        if (diff < 0) diff += 24;
        suma += diff;
      }
    });
    return parseFloat(suma.toFixed(1));
  };

  const estadoLocal = (cp: EmpleadoLocal) => {
    if (faltasHoy[cp.id_ps]) return "Falta";
    if (descansosHoy[cp.id_ps]) return "Descanso manual";
    const t = cp.turnos?.[diaSeleccionada];
    if (!t || t.descanso || t.e === "Descanso") return "Descanso programado";
    const curr = Number(horaSeleccionada.split(":")[0]) * 60 + Number(horaSeleccionada.split(":")[1]);
    const ent = Number(t.e.split(":")[0]) * 60 + Number(t.e.split(":")[1]);
    const sal = Number(t.s.split(":")[0]) * 60 + Number(t.s.split(":")[1]);
    if (curr < ent || curr >= sal) return "Fuera de horario";
    return "Operando en piso";
  };

  const empleadosVista = useMemo(() => {
    if (kpisBackend?.empleados_detallados?.length) return kpisBackend.empleados_detallados.map((e: any) => ({ ...e, nombre: fixText(e.nombre), area: fixText(e.area), estado: fixText(e.estado) }));
    return cinepolitosMatriz.map((cp) => ({
      ...cp,
      id: cp.id_ps,
      area_planificada: cp.area,
      area_excel: cp.area,
      estado: estadoLocal(cp),
      horario: cp.turnos?.[diaSeleccionada]?.descanso ? "Descanso" : `${cp.turnos?.[diaSeleccionada]?.e || ""}-${cp.turnos?.[diaSeleccionada]?.s || ""}`,
      horas_usadas: calcularHorasIndividuales(cp.turnos),
    }));
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [kpisBackend, cinepolitosMatriz, diaSeleccionada, horaSeleccionada, faltasHoy, descansosHoy]);

  const metricas = useMemo(() => {
    if (kpisBackend) {
      const costo = Number(kpisBackend.costo_total_estimado || 0);
      return {
        clientes: Number(kpisBackend.demanda || 0),
        demandaDia: Number(kpisBackend.demanda_dia || 0),
        enPiso: Number(kpisBackend.activos || 0),
        axh: Number(kpisBackend.axh || 0),
        desv: Number(kpisBackend.desviacion_meta || 0),
        horasAsignadasTotales: Number(kpisBackend.horas_asignadas_sem || 0),
        costoNomina: costo,
        cobertura: Number(kpisBackend.cobertura_operativa || 0),
        costoPorCliente: Number(kpisBackend.costo_por_cliente || 0),
      };
    }
    const horaTruncada = `${horaSeleccionada.split(":")[0]}:00`;
    const factor = DISTRIBUCION_DEMANDA[diaSeleccionada]?.[horaTruncada] || 0;
    const clientes = Math.round(demandaSemanaActual * factor);
    const enPiso = empleadosVista.filter((e: any) => e.estado === "Operando en piso" || e.estado === "Reasignado por DSS").length;
    const horas = cinepolitosMatriz.reduce((acc, cp) => acc + calcularHorasIndividuales(cp.turnos), 0);
    const axh = enPiso > 0 ? clientes / enPiso : 0;
    return { clientes, demandaDia: demandaSemanaActual, enPiso, axh, desv: axh - AXH_META, horasAsignadasTotales: horas, costoNomina: horas * COSTO_POR_HORA_PROMEDIO, cobertura: 0, costoPorCliente: 0 };
  }, [kpisBackend, empleadosVista, demandaSemanaActual, diaSeleccionada, horaSeleccionada, cinepolitosMatriz]);

  const areaVisible = (e: any) => fixText(e.area || e.area_planificada || e.area_base || "Lobby");
  const activosPorArea = (area: string) => empleadosVista.filter((e: any) => areaVisible(e) === area && ["Operando en piso", "Reasignado por DSS"].includes(e.estado));
  const rotacion = empleadosVista.filter((e: any) => ["Ley silla", "Alimentos", "Descanso manual", "Descanso DSS"].includes(e.estado));
  const recomendacionesDSS = kpisBackend?.recomendaciones || [];
  const axhSemana = kpisBackend?.axh_semana || [];
  const analytics = kpisBackend?.analytics || {};
  const resumen = kpisBackend?.resumen_dss || {};
  const alertasDSS = (kpisBackend?.alertas || []).map((a: any) => fixText(a));
  const porcDulceria = Number(kpisBackend?.porcentaje_transacciones_dulceria || 0.8052) * 100;
  const disponiblesM2 = Number(kpisBackend?.disponibles_m2_dulceria || 0);
  const planResumen = kpisBackend?.plan_resumen || [];

  const conteoNiveles = [1, 2, 3].map((nivel) => ({ nivel, total: empleadosVista.filter((e: any) => Number(e.nivel_esp ?? e.nivel) === nivel).length }));

  const KpiCard = ({ title, value, sub, tone = "blue" }: { title: string; value: string; sub?: string; tone?: "blue" | "green" | "rose" | "amber" | "slate" }) => {
    const tones: Record<string, string> = {
      blue: "from-blue-50 to-white border-blue-100 text-blue-800",
      green: "from-emerald-50 to-white border-emerald-100 text-emerald-800",
      rose: "from-rose-50 to-white border-rose-100 text-rose-800",
      amber: "from-amber-50 to-white border-amber-100 text-amber-800",
      slate: "from-slate-50 to-white border-slate-100 text-slate-800",
    };
    return (
      <div className={`rounded-2xl border bg-gradient-to-br p-4 shadow-sm ${tones[tone]}`}>
        <p className="text-[10px] font-black uppercase tracking-wider opacity-70">{title}</p>
        <p className="mt-2 text-2xl font-black tracking-tight">{value}</p>
        {sub && <p className="mt-1 text-[10px] font-bold opacity-70">{sub}</p>}
      </div>
    );
  };

  return (
    <div className="flex h-screen bg-[#f8fafc] text-slate-700 font-sans antialiased">
      <aside className="w-64 shrink-0 bg-[#0A192F] text-white p-5 flex flex-col justify-between shadow-xl">
        <div>
          <div className="mb-6 border-b border-slate-700 pb-4">
            <h1 className="text-sm font-black tracking-wider text-[#FFC72C] flex items-center gap-2"><TrendingUp size={18} /> CINEOPS ENGINE</h1>
            <p className="text-[9px] text-slate-400 font-bold tracking-widest uppercase mt-0.5">MILP OPTIMIZER 2026</p>
          </div>

          <div className="mb-4 bg-slate-800/80 p-3 rounded-xl border border-dashed border-slate-600">
            <label className="text-[9px] font-black tracking-wider text-slate-300 uppercase block mb-2 flex items-center gap-1"><FileSpreadsheet size={12} className="text-[#FFC72C]" /> Matriz Horaria</label>
            <div className="relative flex items-center justify-center bg-slate-900 rounded-lg py-3 cursor-pointer hover:bg-slate-950 transition-colors">
              <input type="file" accept=".xlsx,.xls" onChange={handleFileUpload} className="absolute inset-0 opacity-0 cursor-pointer" />
              <div className="text-center text-[11px] text-slate-400 font-medium px-2">
                <Upload size={16} className="mx-auto mb-1 text-[#FFC72C]" />
                {nombreArchivo ? <span className="text-white font-bold truncate block max-w-[180px]">{nombreArchivo}</span> : "Subir matriz de turnos"}
              </div>
            </div>
          </div>

          {estadoCarga === "leyendo" && <div className="mb-4 p-2 rounded-lg bg-amber-500/20 text-amber-300 border border-amber-500/40 text-[10px] font-bold flex items-center gap-2"><Loader2 size={14} className="animate-spin" /> Leyendo filas e integrando demanda...</div>}
          {estadoCarga === "exito" && <div className="mb-4 p-2 rounded-lg bg-emerald-500/20 text-emerald-300 border border-emerald-500/40 text-[10px] font-bold flex items-center gap-2"><CheckCircle2 size={14} /> Se cargaron {cinepolitosMatriz.length} cinepolitos.</div>}
          {estadoCarga === "error" && <div className="mb-4 p-2 rounded-lg bg-rose-500/20 text-rose-300 border border-rose-500/40 text-[10px] font-bold flex items-center gap-2"><AlertCircle size={14} /> Error al mapear la matriz.</div>}

          <nav className="space-y-1 mb-6">
            <button onClick={() => setSubVista("mapa")} className={`w-full p-2.5 rounded-lg flex items-center text-xs font-bold ${subVista === "mapa" ? "bg-[#004B87] text-white" : "text-slate-400 hover:bg-slate-800 hover:text-white"}`}><MapPin className="mr-2.5" size={15} /> Monitoreo en Tiempo Real</button>
            <button onClick={() => setSubVista("plan-semanal")} className={`w-full p-2.5 rounded-lg flex items-center text-xs font-bold ${subVista === "plan-semanal" ? "bg-[#004B87] text-white" : "text-slate-400 hover:bg-slate-800 hover:text-white"}`}><CalendarDays className="mr-2.5" size={15} /> Horarios & Capacidades</button>
            <button onClick={() => setSubVista("analisis")} className={`w-full p-2.5 rounded-lg flex items-center text-xs font-bold ${subVista === "analisis" ? "bg-[#004B87] text-white" : "text-slate-400 hover:bg-slate-800 hover:text-white"}`}><BarChart3 className="mr-2.5" size={15} /> Panel de Control Analítico</button>
            <button onClick={() => setSubVista("asistencia")} className={`w-full p-2.5 rounded-lg flex items-center text-xs font-bold ${subVista === "asistencia" ? "bg-[#004B87] text-white" : "text-slate-400 hover:bg-slate-800 hover:text-white"}`}><ClipboardList className="mr-2.5" size={15} /> Control Asistencias</button>
          </nav>
        </div>

        <div className="space-y-2 bg-slate-900/60 p-3 rounded-xl border border-slate-700 text-[11px]">
          <div className="flex justify-between"><span className="text-slate-400 font-bold">Límite horas:</span><span className="font-mono text-[#FFC72C] font-bold">{presupuestoHorasMax} hrs</span></div>
          <div className="flex justify-between"><span className="text-slate-400 font-bold">Asignado DSS:</span><span className="font-mono text-emerald-400 font-bold">{metricas.horasAsignadasTotales.toFixed(1)} hrs</span></div>
        </div>
      </aside>

      <div className="flex-1 flex flex-col overflow-hidden">
        <header className="bg-white border-b border-slate-200 px-6 py-3 flex items-center justify-between shadow-sm shrink-0">
          <div className="flex items-center gap-3">
            <h2 className="text-xs font-black text-slate-800 uppercase tracking-wider">Módulo DSS: <span className="text-[#004B87]">{subVista.toUpperCase()}</span></h2>
            <div className="flex gap-1.5 items-center bg-slate-100 px-2.5 py-1 rounded-xl border border-slate-200 ml-4">
              <select value={diaSeleccionada} onChange={(e) => setDiaSeleccionada(e.target.value)} className="bg-transparent text-xs font-bold text-slate-700 outline-none cursor-pointer pr-1">
                {DIAS.map((d) => <option key={d} value={d}>{d}</option>)}
              </select>
              <input type="time" value={horaSeleccionada} onChange={(e) => setHoraSeleccionada(e.target.value)} className="bg-transparent text-xs font-bold text-slate-700 outline-none cursor-pointer border-l border-slate-300 pl-1.5" />
            </div>
            <button onClick={() => recalcularOptimizacion(false)} disabled={optimizando} className="rounded-xl bg-[#004B87] text-white px-3 py-1.5 text-[10px] font-black flex items-center gap-1 disabled:opacity-50">
              <RefreshCw size={12} className={optimizando ? "animate-spin" : ""} /> Recalcular Optimización
            </button>
          </div>

          <div className="flex items-center gap-4">
            <div className="text-right border-r pr-3"><span className="block text-[8px] font-bold text-slate-400 uppercase">Costo Est. Nómina</span><span className="text-xs font-black text-emerald-600 flex items-center justify-end gap-0.5"><DollarSign size={12} />{money(metricas.costoNomina)}</span></div>
            <div className="text-right border-r pr-3"><span className="block text-[8px] font-bold text-slate-400 uppercase">AxH Real / Meta</span><span className="text-xs font-black text-slate-800">{metricas.axh.toFixed(1)} / {AXH_META.toFixed(1)}</span></div>
            <div className="text-right border-r pr-3"><span className="block text-[8px] font-bold text-slate-400 uppercase">Desviación</span><span className={`text-xs font-black ${metricas.desv > 0 ? "text-rose-600" : "text-blue-600"}`}>{metricas.desv.toFixed(1)}</span></div>
            <div className="text-right border-r pr-3"><span className="block text-[8px] font-bold text-slate-400 uppercase">Traffic Flow</span><span className="text-xs font-black text-slate-800">{metricas.clientes} adm</span></div>
            <div className="text-right"><span className="block text-[8px] font-bold text-slate-400 uppercase">Activos en Piso</span><span className="text-xs font-black text-[#004B87]">{metricas.enPiso} ch</span></div>
          </div>
        </header>

        <main className="flex-1 p-6 space-y-5 overflow-y-auto">
          {mensajeDSS && <div className="rounded-xl border border-blue-100 bg-blue-50 px-4 py-3 text-[11px] font-bold text-blue-900">{fixText(mensajeDSS)}</div>}

          {subVista === "mapa" && (
            <div className="space-y-5">
              <div className="bg-blue-50 border border-blue-200 rounded-2xl p-4 flex gap-3 items-start shadow-sm">
                <Activity className="text-[#004B87] mt-0.5 shrink-0 animate-pulse" size={18} />
                <div className="space-y-1">
                  <h4 className="text-xs font-black text-blue-900 uppercase tracking-wide">Cuadro de Directivas Operativas ({diaSeleccionada} {horaSeleccionada})</h4>
                  <div className="text-[11px] text-blue-800 leading-relaxed space-y-1">
                    {metricas.axh > 7.5 ? (
                      <p><strong>Alerta de sub-staffing:</strong> AxH en <b className="text-rose-600">{metricas.axh.toFixed(1)}</b>. Conviene reforzar áreas críticas, especialmente Dulcería.</p>
                    ) : (
                      <p><strong>Operación controlada:</strong> el nivel de atención se mantiene cerca de la meta. Supervisa Ley Silla y alimentos escalonados.</p>
                    )}
                    {alertasDSS.slice(0, 3).map((a: string, i: number) => (
                      <p key={i} className="font-bold text-amber-800">⚠ {a}</p>
                    ))}
                  </div>
                </div>
              </div>

              <section className="space-y-2">
                <h3 className="text-xs font-black text-slate-800 uppercase tracking-wider flex items-center gap-1.5 border-b pb-1"><span className="w-2 h-2 rounded-full bg-blue-600"></span>1. Distribución Activa en Estaciones de Trabajo</h3>
                <div className="grid grid-cols-5 gap-3">
                  {AREAS.map((area) => {
                    const activos = activosPorArea(area);
                    return (
                      <div key={area} className="bg-white border border-slate-200 rounded-2xl p-3 shadow-sm min-h-[235px]">
                        <div className="flex justify-between items-center border-b pb-2 mb-2">
                          <span className="text-[11px] font-black text-slate-800 uppercase">{area}</span>
                          <span className="bg-[#004B87] text-white font-mono px-2 py-0.5 rounded text-[9px] font-bold">{activos.length}</span>
                        </div>
                        <div className="space-y-1.5 max-h-[185px] overflow-y-auto pr-1">
                          {activos.length === 0 ? <p className="text-[10px] text-slate-400 italic py-8 text-center">Sin personal activo</p> : activos.map((cp: any) => (
                            <div key={idOf(cp)} title={fixText(cp.nombre)} className="p-2 bg-slate-50 border border-slate-100 rounded-xl text-[10px] font-bold text-slate-800 hover:border-blue-200 hover:bg-blue-50 transition-colors">
                              <div className="truncate max-w-[185px]">{fixText(cp.nombre)}</div>
                              <div className="mt-1 flex gap-1 flex-wrap items-center">
                                <span className="text-[8px] text-slate-500 font-mono">PS {idOf(cp)}</span>
                                <span className="text-[8px] bg-slate-200 text-slate-700 px-1 rounded">N{cp.nivel_esp ?? cp.nivel}</span>
                                <span className={`text-[8px] px-1 rounded border ${estadoColor(cp.estado)}`}>{fixText(cp.estado)}</span>
                              </div>
                            </div>
                          ))}
                        </div>
                      </div>
                    );
                  })}
                </div>
              </section>

              <div className="grid grid-cols-2 gap-4">
                <section className="bg-amber-50/60 border border-amber-200 rounded-2xl p-4 space-y-3">
                  <h3 className="text-xs font-black text-amber-800 uppercase tracking-wider flex items-center gap-1.5 border-b border-amber-200 pb-1.5"><span className="w-2 h-2 rounded-full bg-amber-500"></span>2. Personal en Rotación Activa</h3>
                  <div className="grid grid-cols-2 gap-2 max-h-[210px] overflow-y-auto pr-1">
                    {rotacion.length === 0 ? <p className="text-[10px] text-slate-400 italic text-center py-8 col-span-2">Sin Ley Silla, alimentos o descansos activos en esta hora.</p> : rotacion.map((cp: any) => (
                      <div key={idOf(cp)} className="bg-white p-2.5 rounded-xl border border-amber-200 text-[10px] space-y-1">
                        <div className="font-black text-slate-900 truncate" title={fixText(cp.nombre)}>{fixText(cp.nombre)}</div>
                        <div className="flex justify-between items-center gap-2">
                          <span className={`px-1.5 py-0.5 rounded text-[8px] font-black uppercase border ${estadoColor(cp.estado)}`}>{fixText(cp.estado)}</span>
                          <span className="text-[8px] font-mono text-slate-500 font-bold truncate">{fixText(cp.detalle_break || cp.horario || "")}</span>
                        </div>
                      </div>
                    ))}
                  </div>
                </section>

                <section className="bg-blue-50/70 border border-blue-200 rounded-2xl p-4 space-y-3">
                  <h3 className="text-xs font-black text-blue-900 uppercase tracking-wider flex items-center gap-1.5 border-b border-blue-200 pb-1.5"><span className="w-2 h-2 rounded-full bg-blue-600"></span>3. Recomendaciones DSS para Cubrir Faltas</h3>
                  <div className="flex items-center justify-between gap-2">
                    <p className="text-[10px] text-blue-800 leading-normal">Ordena candidatos por disponibilidad, horas restantes, área crítica y polivalencia M1-M5.</p>
                    <div className="flex gap-1 items-center">
                      <select value={descansoAutoDias} onChange={(e) => setDescansoAutoDias(Number(e.target.value) as 1 | 2)} className="text-[10px] bg-white border border-blue-200 rounded px-1 py-1 font-bold">
                        <option value={1}>1 día descanso</option>
                        <option value={2}>2 días descanso</option>
                      </select>
                      <button onClick={() => recalcularOptimizacion(true)} disabled={optimizando} className="bg-blue-600 text-white text-[9px] font-black rounded px-2 py-1 disabled:opacity-60">AUTO</button>
                      <button onClick={() => recalcularOptimizacion(false)} disabled={optimizando} className="bg-[#004B87] text-white text-[9px] font-black rounded px-2 py-1 disabled:opacity-60">RECALCULAR</button>
                    </div>
                  </div>
                  <div className="max-h-[210px] overflow-y-auto space-y-2 pr-1">
                    {recomendacionesDSS.length === 0 ? <div className="p-3 bg-white rounded-xl border border-blue-100 text-[10px] text-slate-500 font-bold text-center">Sin faltas críticas. Al marcar una falta aparecerá el mejor reemplazo sugerido.</div> : recomendacionesDSS.slice(0, 8).map((r: any, idx: number) => (
                      <div key={`${r.id}-${idx}`} className="bg-white p-2.5 rounded-xl border border-blue-100 text-[10px] shadow-sm">
                        <div className="flex justify-between gap-2">
                          <span className="font-black text-slate-900 truncate" title={fixText(r.empleado)}>{idx + 1}. {fixText(r.empleado)}</span>
                          <span className="font-mono text-[8px] text-blue-700 bg-blue-100 px-1 rounded">Score {r.score}</span>
                        </div>
                        <div className="flex justify-between mt-1 text-[9px] text-slate-500 font-bold"><span>{fixText(r.origen || "Disponible")} → <b className="text-blue-800">{fixText(r.mover_a)}</b></span><span>{r.horas_restantes} hrs</span></div>
                        <p className="text-[9px] text-slate-500 mt-1 leading-normal">{fixText(r.motivo)}</p>
                        <button onClick={() => aplicarRecomendacion(r)} className="mt-2 w-full rounded-lg bg-emerald-600 text-white py-1 text-[9px] font-black">APLICAR RECOMENDACIÓN</button>
                      </div>
                    ))}
                  </div>
                </section>
              </div>
            </div>
          )}

          {subVista === "plan-semanal" && (
            <div className="bg-white rounded-2xl border border-slate-200 overflow-hidden shadow-sm">
              <div className="p-4 border-b bg-slate-50 flex items-center justify-between">
                <div>
                  <h3 className="text-xs font-black uppercase text-slate-800">Plan Semanal DSS</h3>
                  <p className="text-[10px] text-slate-500 font-bold">Área y horario asignados por el optimizador, sin saturar la vista con etiquetas internas.</p>
                </div>
                <button onClick={() => recalcularOptimizacion(false)} className="rounded-lg bg-[#004B87] text-white px-3 py-1.5 text-[10px] font-black">Recalcular Optimización</button>
              </div>
              <div className="overflow-x-auto max-h-[calc(100vh-190px)]">
                <table className="w-full text-left border-collapse text-[11px]">
                  <thead className="bg-slate-900 text-slate-300 text-[9px] uppercase font-black sticky top-0 z-10">
                    <tr>
                      <th className="p-3 sticky left-0 bg-slate-900 text-white min-w-[220px]">Cinepolito</th>
                      <th className="p-2 text-center bg-slate-800">Nivel</th>
                      <th className="p-2 text-center text-[#FFC72C] border-r border-slate-700">M1-M5</th>
                      {DIAS.map((d) => <th key={d} className="p-2 text-center min-w-[145px]">{d}</th>)}
                      <th className="p-3 text-center bg-[#004B87] text-white">Horas</th>
                    </tr>
                  </thead>
                  <tbody>
                    {(planResumen.length ? planResumen : empleadosVista.map((e: any) => ({
                      id: idOf(e),
                      nombre: e.nombre,
                      nivel: e.nivel_esp,
                      habilidades: e.expertise,
                      dias: DIAS.map((d) => ({
                        dia: d,
                        dss: e.area || e.area_planificada || "Lobby",
                        horario: e.horario || "Horario asignado",
                      })),
                      horas_semana: e.horas_usadas || 0,
                    }))).map((row: any) => (
                      <tr key={row.id} className="hover:bg-slate-50 border-b border-slate-200">
                        <td className="p-2 font-bold text-slate-800 sticky left-0 bg-white z-[1]">
                          <div className="truncate max-w-[210px]" title={fixText(row.nombre)}>{fixText(row.nombre)}</div>
                          <div className="text-[9px] font-mono text-slate-400">PS {row.id}</div>
                        </td>
                        <td className="p-2 text-center font-black text-blue-700">{row.nivel ?? row.nivel_esp}⭐</td>
                        <td className="p-2 text-center border-r border-slate-200">
                          <div className="flex gap-0.5 justify-center flex-wrap">
                            {Object.entries(row.habilidades || {}).map(([k, v]) => <span key={k} className={`text-[8px] font-mono px-1 rounded font-bold ${Number(v) === 1 ? "bg-emerald-100 text-emerald-800 border border-emerald-300" : "bg-slate-100 text-slate-300"}`}>{k}:{Number(v)}</span>)}
                          </div>
                        </td>
                        {row.dias.map((d: any) => {
                          const area = fixText(d.area || d.dss || "Descanso");
                          const horario = fixText(d.horario || (area === "Descanso" ? "Descanso" : "Horario asignado"));
                          return (
                            <td key={d.dia} className={`p-2 text-center text-[10px] font-bold border-l ${areaCellColor(area)}`} title={`${area} · ${horario}`}>
                              <div className="text-[11px] font-black">{area}</div>
                              <div className="text-[9px] font-mono opacity-70 mt-1">{area === "Descanso" ? "" : horario}</div>
                            </td>
                          );
                        })}
                        <td className="p-3 text-center font-mono font-black bg-blue-50 text-[#004B87]">{Number(row.horas_semana || 0).toFixed(0)} hrs</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          )}

          {subVista === "analisis" && (
            <div className="space-y-5">
              <div className="bg-white p-5 rounded-2xl border border-slate-200 shadow-sm flex items-center gap-3">
                <div className="bg-[#004B87] text-white p-3 rounded-xl"><BarChart3 size={22} /></div>
                <div><h3 className="text-sm font-black uppercase text-slate-800">Dashboard Ejecutivo de Control Analítico</h3><p className="text-[11px] text-slate-500 font-bold">KPIs operativos, financieros y de servicio de la solución DSS.</p></div>
              </div>

              {alertasDSS.length > 0 && (
                <div className="bg-amber-50 border border-amber-200 rounded-2xl p-4 shadow-sm">
                  <h4 className="text-[10px] font-black uppercase text-amber-800 tracking-wider mb-2">Alertas ejecutivas DSS</h4>
                  <div className="grid grid-cols-1 gap-2">
                    {alertasDSS.slice(0, 4).map((a: string, i: number) => (
                      <div key={i} className="bg-white/70 border border-amber-100 rounded-xl px-3 py-2 text-[10px] font-bold text-amber-900">⚠ {a}</div>
                    ))}
                  </div>
                </div>
              )}

              <div className="grid grid-cols-6 gap-3">
                <KpiCard title="AxH Real / Meta" value={`${metricas.axh.toFixed(1)} / 7.0`} sub="Productividad de atención" tone={metricas.axh > 8 ? "rose" : "green"} />
                <KpiCard title="Demanda Hora" value={`${metricas.clientes}`} sub={`${diaSeleccionada} ${horaSeleccionada}`} tone="blue" />
                <KpiCard title="Activos en Piso" value={`${metricas.enPiso}`} sub={`${metricas.cobertura.toFixed(1)}% cobertura`} tone="green" />
                <KpiCard title="Consumo Horas" value={`${Number(kpisBackend?.consumo_horas_pct || (metricas.horasAsignadasTotales / presupuestoHorasMax) * 100).toFixed(1)}%`} sub={`${metricas.horasAsignadasTotales.toFixed(0)} / ${presupuestoHorasMax}`} tone={Number(kpisBackend?.consumo_horas_pct || 0) >= 95 ? "rose" : "amber"} />
                <KpiCard title="Dulcería" value={`${porcDulceria.toFixed(2)}%`} sub="Transacciones 2025" tone="green" />
                <KpiCard title="M2 disponibles" value={`${disponiblesM2}`} sub="Capacitados Dulcería" tone={disponiblesM2 < Number((analytics.personal_ideal || [])[1] || 0) ? "rose" : "green"} />
              </div>

              <div className="grid grid-cols-2 gap-4">
                <div className="bg-white p-4 rounded-2xl border border-slate-200 shadow-sm">
                  <h4 className="text-[10px] font-black uppercase text-slate-500 tracking-wider mb-3">Cobertura, déficit y exceso por área</h4>
                  {(analytics.labels_areas || AREAS).map((area: string, idx: number) => {
                    const real = Number(analytics.personal_real?.[idx] || 0);
                    const ideal = Number(analytics.personal_ideal?.[idx] || 0);
                    const deficit = Number(analytics.deficit_area?.[idx] || 0);
                    const exceso = Number(analytics.exceso_area?.[idx] || 0);
                    const max = Math.max(1, real, ideal);
                    return <div key={area} className="mb-3"><div className="flex justify-between text-[10px] font-black text-slate-600"><span>{fixText(area)}</span><span>Real {real} / Ideal {ideal} · Déficit {deficit} · Exceso {exceso}</span></div><div className="h-3 bg-slate-100 rounded-full overflow-hidden mt-1 flex"><div className="bg-[#004B87]" style={{ width: barWidth(real, max) }}></div><div className="bg-slate-300" style={{ width: barWidth(ideal, max) }}></div></div></div>;
                  })}
                </div>

                <div className="bg-white p-4 rounded-2xl border border-slate-200 shadow-sm">
                  <h4 className="text-[10px] font-black uppercase text-slate-500 tracking-wider mb-3">AxH por día: solución DSS vs meta</h4>
                  {(axhSemana.length ? axhSemana : DIAS.map((d) => ({ dia: d, axh_optimo: metricas.axh, demanda: 0 }))).map((d: any) => {
                    const axhDss = Number(d.axh_optimo || d.axh || 0);
                    const max = Math.max(12, axhDss, 7);
                    return <div key={d.dia} className="mb-3"><div className="flex justify-between text-[10px] font-black text-slate-600"><span>{fixText(d.dia)}</span><span>DSS {axhDss.toFixed(1)} / Meta 7.0</span></div><div className="grid grid-cols-1 gap-1 mt-1"><div className="h-3 bg-blue-100 rounded"><div className="h-3 bg-[#004B87] rounded" style={{ width: barWidth(axhDss, max) }}></div></div><div className="h-1 bg-emerald-100 rounded"><div className="h-1 bg-emerald-500 rounded" style={{ width: barWidth(7, max) }}></div></div></div></div>;
                  })}
                  <div className="flex gap-3 text-[9px] font-bold text-slate-500"><span><b className="text-[#004B87]">■</b> DSS</span><span><b className="text-emerald-500">■</b> Meta</span></div>
                </div>
              </div>

              <div className="grid grid-cols-3 gap-3">
                <KpiCard title="Demanda diaria" value={`${metricas.demandaDia}`} sub="Admisiones del día" tone="blue" />
                <KpiCard title="Polivalencia" value={`${Number(resumen.polivalencia_promedio || 0).toFixed(2)}`} sub={`${resumen.empleados_multi_area || 0} multiárea`} tone="green" />
                <KpiCard title="Faltas / Descansos" value={`${resumen.faltas || 0} / ${resumen.descansos || 0}`} sub="Control operativo" tone="amber" />
              </div>
            </div>
          )}

          {subVista === "asistencia" && (
            <div className="space-y-6">
              <div className="bg-white p-3.5 rounded-2xl border border-slate-200 text-xs font-bold text-slate-500 flex justify-between items-center">
                <span>Pasa lista por área. “En piso” registra falta; “Descanso” registra descanso manual sin mezclarlo con falta.</span>
                <button onClick={() => recalcularOptimizacion(false)} className="rounded-lg bg-[#004B87] text-white px-3 py-1.5 text-[10px] font-black">Recalcular Optimización</button>
              </div>
              {AREAS.map((area) => {
                const cinepolitosDelArea = empleadosVista.filter((c: any) => areaVisible(c) === area);
                if (cinepolitosDelArea.length === 0) return null;
                return (
                  <div key={area} className="bg-white rounded-2xl border border-slate-200 shadow-sm overflow-hidden">
                    <div className="bg-slate-900 text-white px-4 py-2 flex justify-between items-center"><span className="text-xs font-black uppercase tracking-wider">{area}</span><span className="bg-[#004B87] text-white px-2 py-0.5 rounded text-[10px] font-bold">{cinepolitosDelArea.length} asignados</span></div>
                    <table className="w-full text-left border-collapse text-[11px]">
                      <thead><tr className="bg-slate-50 text-slate-500 font-bold text-[9px] uppercase border-b"><th className="p-2.5 pl-4 w-28">ID PS</th><th className="p-2.5">Colaborador</th><th className="p-2.5 text-center w-56">Acciones</th></tr></thead>
                      <tbody className="divide-y divide-slate-100">
                        {cinepolitosDelArea.map((c: any) => {
                          const id = idOf(c);
                          const tieneFalta = faltasHoy[id] || c.estado === "Falta";
                          const tieneDescanso = descansosHoy[id] || String(c.estado).includes("Descanso manual");
                          return <tr key={id} className={`hover:bg-slate-50/60 transition-colors ${tieneFalta ? "bg-rose-50" : tieneDescanso ? "bg-sky-50" : ""}`}><td className="p-2.5 pl-4 font-mono font-bold text-slate-500">{id}</td><td className={`p-2.5 font-bold ${tieneFalta ? "line-through text-slate-400" : "text-slate-900"}`} title={fixText(c.nombre)}>{fixText(c.nombre)}</td><td className="p-2 text-center"><button onClick={() => marcarFaltaBackend(c)} className={`px-3 py-1 rounded font-black text-[9px] tracking-wide border transition-all ${tieneFalta ? "bg-rose-600 text-white border-rose-700" : "bg-white text-emerald-700 border-emerald-200 hover:bg-emerald-50"}`}>{tieneFalta ? "FALTA" : "EN PISO"}</button><button onClick={() => marcarDescansoManualBackend(c)} className={`ml-1 px-3 py-1 rounded font-black text-[9px] tracking-wide border transition-all ${tieneDescanso ? "bg-sky-600 text-white border-sky-700" : "bg-amber-50 text-amber-700 border-amber-200 hover:bg-amber-100"}`}>{tieneDescanso ? "DESCANSO" : "DESCANSO"}</button></td></tr>;
                        })}
                      </tbody>
                    </table>
                  </div>
                );
              })}
            </div>
          )}
        </main>
      </div>
    </div>
  );
}
