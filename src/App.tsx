import { useState, useMemo, useEffect } from 'react';
import { motion, AnimatePresence } from 'framer-motion';
import { 
  Users, 
  Droplets, 
  Clock, 
  Calculator, 
  Plus, 
  Trash2, 
  Calendar as CalendarIcon,
  ChevronRight,
  Info,
  AlertCircle,
  FileDown,
  Printer
} from 'lucide-react';
import * as XLSX from 'xlsx';

interface Farmer {
  id: number;
  order: number;
  duration: number; // in hours
  flow: number;     // Q in L/s
}

interface ScheduledFarmer extends Farmer {
  startTime: number; // in hours from start of tour
  endTime: number;
  volume: number;    // V in m3
}

export default function App() {
  const [farmerCount, setFarmerCount] = useState<number>(5);
  const [defaultMaxFlow, setDefaultMaxFlow] = useState<number>(50); // Default L/s
  const [dailyMaxFlows, setDailyMaxFlows] = useState<Record<number, number>>({}); // Day Index -> Max Flow L/s
  const [tourDays, setTourDays] = useState<number>(7);
  const [farmers, setFarmers] = useState<Farmer[]>([]);
  const [scheduledResults, setScheduledResults] = useState<ScheduledFarmer[]>([]);
  const [tourStartDate, setTourStartDate] = useState<string>(new Date().toISOString().split('T')[0]);
  const [isCalculating, setIsCalculating] = useState(false);
  const [isDirty, setIsDirty] = useState(false);
  const [activeTab, setActiveTab] = useState<'dashboard' | 'recap'>('dashboard');

  // Persistence: Load on mount
  useEffect(() => {
    const saved = localStorage.getItem('irrigaPlan_data');
    if (saved) {
      try {
        const data = JSON.parse(saved);
        setFarmerCount(data.farmerCount || 5);
        setDefaultMaxFlow(data.defaultMaxFlow || 50);
        setDailyMaxFlows(data.dailyMaxFlows || {});
        setTourDays(data.tourDays || 7);
        setFarmers(data.farmers || []);
        setTourStartDate(data.tourStartDate || new Date().toISOString().split('T')[0]);
      } catch (e) {
        console.error("Failed to load saved data", e);
      }
    }
  }, []);

  // Persistence: Save on change
  useEffect(() => {
    const data = {
      farmerCount,
      defaultMaxFlow,
      dailyMaxFlows,
      tourDays,
      farmers,
      tourStartDate
    };
    localStorage.setItem('irrigaPlan_data', JSON.stringify(data));
  }, [farmerCount, defaultMaxFlow, dailyMaxFlows, tourDays, farmers, tourStartDate]);

  const handlePrint = () => {
    window.print();
  };

  const handleExportExcel = () => {
    if (scheduledResults.length === 0) return;

    const data = scheduledResults.map((f, idx) => ({
      'N° Ordre': idx + 1,
      'Durée (h)': f.duration,
      'Débit (L/S)': f.flow,
      'Volume (m³)': parseFloat((f.flow * f.duration * 3.6).toFixed(2)),
      'Début': formatTime(f.startTime),
      'Fin': formatTime(f.endTime)
    }));

    const worksheet = XLSX.utils.json_to_sheet(data);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Plan Irrigation");

    // Summary calculation for each day
    const daySummary = Array.from({ length: effectiveTourDays }).map((_, dayIdx) => {
      const dayFarmers = scheduledResults.filter(f => 
        getDayIndex(f.startTime) === dayIdx || 
        getDayIndex(f.endTime) === dayIdx
      );

      const dayStart = (dayIdx * 24) - 9;
      const dayEnd = ((dayIdx + 1) * 24) - 9;
      
      const dayTotalVolume = dayFarmers.reduce((sum, f) => {
        const actualStart = Math.max(f.startTime, dayStart);
        const actualEnd = Math.min(f.endTime, dayEnd);
        const durationInDay = Math.max(0, actualEnd - actualStart);
        return sum + (f.flow * durationInDay * 3.6);
      }, 0);

      return {
        'Jour': dayIdx + 1,
        'Date': new Date(new Date(tourStartDate).getTime() + dayIdx * 24 * 60 * 60 * 1000).toLocaleDateString('fr-FR'),
        'Volume Total (m³)': parseFloat(dayTotalVolume.toFixed(2))
      };
    });

    const summarySheet = XLSX.utils.json_to_sheet(daySummary);
    XLSX.utils.book_append_sheet(workbook, summarySheet, "Récapitulatif Journalier");

    XLSX.writeFile(workbook, `Plan_Irrigation_${tourStartDate}.xlsx`);
  };

  // Initialize farmers when count changes
  useEffect(() => {
    setFarmers(prev => {
      const currentCount = prev.length;
      if (farmerCount > currentCount) {
        const added = Array.from({ length: farmerCount - currentCount }, (_, i) => ({
          id: currentCount + i + 1,
          order: currentCount + i + 1,
          duration: 2,
          flow: 10,
        }));
        return [...prev, ...added];
      } else if (farmerCount < currentCount) {
        return prev.slice(0, farmerCount);
      }
      return prev;
    });
    // Don't reset results here if we want to keep them, but let's mark it as dirty
    setIsDirty(true);
  }, [farmerCount]);

  const updateFarmer = (id: number, field: keyof Farmer, value: string | number) => {
    setFarmers(prev => prev.map(f => f.id === id ? { ...f, [field]: value } : f));
    setIsDirty(true);
  };

  const updateDailyFlow = (day: number, value: number) => {
    setDailyMaxFlows(prev => ({ ...prev, [day]: value }));
  };

  const getDayIndex = (hours: number) => {
    // Current time is hours + 9h (start hour)
    return Math.floor((hours + 9) / 24);
  };

  const getDayMaxFlow = (dayIndex: number) => {
    return dailyMaxFlows[dayIndex] || defaultMaxFlow;
  };

  // Scheduling logic
  const handleCalculate = () => {
    setIsCalculating(true);
    
    setTimeout(() => {
      const sorted = [...farmers].sort((a, b) => a.order - b.order);
      const scheduled: ScheduledFarmer[] = [];
      let events: { time: number; delta: number }[] = [];

      const getFlowAt = (t: number) => {
        let flow = 0;
        for (const event of events) {
          if (event.time <= t + 0.0001) flow += event.delta;
          else break;
        }
        return flow;
      };

      const canFit = (start: number, duration: number, flow: number) => {
        const end = start + duration;
        
        // Day boundaries are when (hours + 9) is a multiple of 24
        // (t + 9) = 24k => t = 24k - 9
        const dayBoundaries = Array.from({length: effectiveTourDays + 2}).map((_, i) => (i + 1) * 24 - 9).filter(t => t > 0);
        
        const checkPoints = [
          start, 
          end, 
          ...events.map(e => e.time).filter(t => t > start && t < end),
          ...dayBoundaries.filter(t => t > start && t < end)
        ].sort((a, b) => a - b);
        
        const uniqueCheckPoints = checkPoints.filter((v, i, a) => i === 0 || v > a[i-1] + 0.0001);

        for (const t of uniqueCheckPoints) {
          const dayIndex = getDayIndex(t);
          const limit = getDayMaxFlow(dayIndex);
          if (flow > limit) return false;
          if (getFlowAt(t) + flow > limit + 0.0001) return false;
        }
        return true;
      };

      for (const f of sorted) {
        let startTime = 0;
        const dayBoundaries = Array.from({length: effectiveTourDays + 2}).map((_, i) => (i + 1) * 24 - 9).filter(t => t >= 0);
        const possibleStarts = [0, ...events.map(e => e.time), ...dayBoundaries]
          .filter(t => t >= 0)
          .sort((a, b) => a - b);
        
        const uniquePossibleStarts = possibleStarts.filter((v, i, a) => i === 0 || v > a[i-1] + 0.0001);

        for (const s of uniquePossibleStarts) {
          if (canFit(s, f.duration, f.flow)) {
            startTime = s;
            break;
          }
        }

        const endTime = startTime + f.duration;
        scheduled.push({
          ...f,
          startTime,
          endTime,
          volume: f.flow * f.duration * 3.6
        });

        events.push({ time: startTime, delta: f.flow });
        events.push({ time: endTime, delta: -f.flow });
        events.sort((a, b) => a.time - b.time || a.delta - b.delta);
      }

      setScheduledResults(scheduled);
      setIsCalculating(false);
      setIsDirty(false);
    }, 600);
  };

  const totalVolume = farmers.reduce((acc, curr) => acc + (curr.flow * curr.duration * 3.6), 0);
  const maxDayReached = Math.max(0, ...scheduledResults.map(f => getDayIndex(f.endTime)));
  const effectiveTourDays = Math.max(tourDays, maxDayReached + 1);

  const formatTime = (hours: number) => {
    const baseDate = new Date(tourStartDate);
    baseDate.setHours(9, 0, 0, 0); 
    const date = new Date(baseDate.getTime() + hours * 60 * 60 * 1000);
    
    // Day label: J1 = first calendar day, etc.
    const dayOfTour = getDayIndex(hours) + 1;
    const timeStr = date.toLocaleTimeString('fr-FR', {
      hour: '2-digit',
      minute: '2-digit'
    });

    return `J${dayOfTour} ${timeStr}`;
  };

  return (
    <div className="min-h-screen bg-[#F8F9FA] text-[#1A1A1A] font-sans selection:bg-blue-100 p-4 md:p-8">
      <div className="max-w-7xl mx-auto space-y-8">
        
        {/* Header */}
        <header className="flex flex-col md:flex-row md:items-end justify-between gap-6 border-b border-gray-200 pb-8 no-print">
          <div className="space-y-1 text-center md:text-left">
            <h1 className="text-4xl font-semibold tracking-tight text-gray-900">
              IrrigaPlan <span className="text-gray-400 font-light">/ {activeTab === 'dashboard' ? 'Tableau' : 'Récapitulatif'}</span>
            </h1>
            <p className="text-gray-500 max-w-sm">
              Gestion des débits et horaires. Début du tour fixé à 09:00 Janvier 1er ou date choisie.
            </p>
          </div>

          <div className="flex gap-1 p-1 bg-gray-100 rounded-xl no-print">
            <button 
              onClick={() => setActiveTab('dashboard')}
              className={`px-4 py-2 rounded-lg text-sm font-medium transition-all ${activeTab === 'dashboard' ? 'bg-white shadow-sm text-blue-600' : 'text-gray-500 hover:text-gray-700'}`}
            >
              Tableau Saisie
            </button>
            <button 
              onClick={() => setActiveTab('recap')}
              className={`px-4 py-2 rounded-lg text-sm font-medium transition-all ${activeTab === 'recap' ? 'bg-white shadow-sm text-blue-600' : 'text-gray-500 hover:text-gray-700'}`}
            >
              Récapitulatif
            </button>
          </div>
        </header>

        <section className="no-print">
          <div className="flex flex-wrap gap-4 items-end">
            <div className="flex flex-col gap-1">
              <label className="text-[10px] font-bold uppercase tracking-widest text-gray-400">Nombre d'agriculteurs</label>
              <div className="relative">
                <Users size={16} className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-400" />
                <input 
                  type="number" 
                  min="1" 
                  max="100"
                  value={farmerCount}
                  onChange={(e) => setFarmerCount(Math.max(1, parseInt(e.target.value) || 1))}
                  className="pl-10 pr-4 py-2 bg-white border border-gray-200 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-transparent outline-none transition-all w-32 font-mono"
                />
              </div>
            </div>

            <div className="flex flex-col gap-1">
              <label className="text-[10px] font-bold uppercase tracking-widest text-gray-400">Débit Max Défaut (L/S)</label>
              <div className="relative">
                <AlertCircle size={16} className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-400" />
                <input 
                  type="number" 
                  min="1" 
                  value={defaultMaxFlow}
                  onChange={(e) => setDefaultMaxFlow(Math.max(1, parseInt(e.target.value) || 1))}
                  className="pl-10 pr-4 py-2 bg-white border border-gray-200 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-transparent outline-none transition-all w-40 font-mono"
                />
              </div>
            </div>
            <div className="flex flex-col gap-1">
              <label className="text-[10px] font-bold uppercase tracking-widest text-gray-400">Nb Jours Tour</label>
              <div className="relative">
                <CalendarIcon size={16} className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-400" />
                <input 
                  type="number" 
                  min="1"
                  max="31"
                  value={tourDays}
                  onChange={(e) => setTourDays(Math.max(1, parseInt(e.target.value) || 1))}
                  className="pl-10 pr-4 py-2 bg-white border border-gray-200 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-transparent outline-none transition-all w-24 font-mono text-sm"
                />
              </div>
            </div>

            <div className="flex flex-col gap-1">
              <input 
                type="date" 
                value={tourStartDate}
                onChange={(e) => setTourStartDate(e.target.value)}
                className="px-4 py-2 bg-white border border-gray-200 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-transparent outline-none transition-all w-44 font-mono text-sm"
              />
            </div>

            <div className="flex flex-col gap-1 no-print">
              <label className="text-[10px] font-bold uppercase tracking-widest text-gray-400 opacity-0">Actions</label>
              <div className="flex gap-2">
                <button 
                  onClick={handleCalculate}
                  disabled={isCalculating}
                  className="flex items-center gap-2 px-6 py-2 bg-blue-600 text-white rounded-lg hover:bg-blue-700 disabled:opacity-50 transition-all font-semibold text-sm shadow-sm active:scale-95"
                >
                  {isCalculating ? "Calcul..." : "Lancer le Tour"}
                </button>
                <button 
                  onClick={handleExportExcel}
                  disabled={scheduledResults.length === 0}
                  className="flex items-center gap-2 px-4 py-2 bg-green-600 text-white rounded-lg hover:bg-green-700 disabled:opacity-50 transition-all font-medium text-sm shadow-sm active:scale-95"
                >
                  <FileDown size={16} />
                  Excel
                </button>
                <button 
                  onClick={handlePrint}
                  className="flex items-center gap-2 px-4 py-2 bg-gray-900 text-white rounded-lg hover:bg-gray-800 transition-all font-medium text-sm shadow-sm active:scale-95"
                >
                  <Printer size={16} />
                  PDF
                </button>
              </div>
            </div>
          </div>
        </section>

        {/* Main Body */}
        {activeTab === 'dashboard' ? (
          <div className="grid grid-cols-1 lg:grid-cols-4 gap-8">
            
            {/* Daily flow limits sidebar */}
            <div className="lg:col-span-1 space-y-4 no-print">
              <div className="bg-white p-6 rounded-2xl border border-gray-100 shadow-sm">
                <h3 className="text-sm font-bold uppercase tracking-widest text-gray-900 mb-4 flex items-center gap-2">
                  <Calculator size={16} className="text-blue-600" />
                  Débit Max Journalier (L/S)
                </h3>
                <div className="space-y-3 max-h-[500px] overflow-y-auto pr-2 custom-scrollbar">
                  {Array.from({ length: effectiveTourDays }).map((_, i) => (
                    <div key={i} className="flex items-center justify-between gap-4 p-2 rounded-lg hover:bg-gray-50 transition-colors">
                      <div className="flex flex-col">
                        <span className="text-[10px] font-bold text-gray-400 uppercase">Jour {i + 1}</span>
                        <span className="text-[10px] text-gray-300">
                          {new Date(new Date(tourStartDate).getTime() + i * 24 * 60 * 60 * 1000).toLocaleDateString('fr-FR', { day: 'numeric', month: 'short' })}
                        </span>
                      </div>
                      <input 
                        type="number"
                        placeholder={defaultMaxFlow.toString()}
                        value={dailyMaxFlows[i] || ''}
                        onChange={(e) => updateDailyFlow(i, parseInt(e.target.value) || 0)}
                        className="w-20 bg-white border border-gray-100 rounded px-2 py-1.5 text-xs font-mono outline-none focus:ring-2 focus:ring-blue-100"
                      />
                    </div>
                  ))}
                </div>
                <p className="text-[10px] text-gray-400 mt-4 leading-tight italic">
                  * Si vide, utilise le débit par défaut ({defaultMaxFlow} L/S).
                </p>
              </div>
            </div>

            <div className="lg:col-span-3 bg-white rounded-2xl shadow-sm border border-gray-100 overflow-hidden print:border-none print:shadow-none min-h-[400px] flex flex-col">
              <div className="overflow-x-auto flex-grow">
                <table className="w-full text-left border-collapse">
                  <thead>
                    <tr className="bg-gray-50/50 border-bottom border-gray-100">
                      <th className="px-6 py-4 text-[10px] font-bold uppercase tracking-widest text-gray-400">N° Ordre</th>
                      <th className="px-6 py-4 text-[10px] font-bold uppercase tracking-widest text-gray-400">Durée (h)</th>
                      <th className="px-6 py-4 text-[10px] font-bold uppercase tracking-widest text-gray-400">Débit (L/S)</th>
                      <th className="px-6 py-4 text-[10px] font-bold uppercase tracking-widest text-gray-400">Volume (m³)</th>
                      <th className="px-6 py-4 text-[10px] font-bold uppercase tracking-widest text-gray-400">Début</th>
                      <th className="px-6 py-4 text-[10px] font-bold uppercase tracking-widest text-gray-400">Fin</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-gray-50">
                    <AnimatePresence mode="popLayout">
                      {farmers.map((farmer, idx) => {
                        const schedule = scheduledResults.find(s => s.id === farmer.id);
                        return (
                          <motion.tr 
                            layout
                            initial={{ opacity: 0, y: 10 }}
                            animate={{ opacity: 1, y: 0 }}
                            exit={{ opacity: 0 }}
                            key={farmer.id}
                            className="group hover:bg-blue-50/30 transition-colors"
                          >
                            <td className="px-6 py-4 font-mono text-sm text-gray-500">
                              {idx + 1}
                            </td>
                            <td className="px-6 py-4">
                              <div className="flex items-center gap-2">
                                <input 
                                  type="number"
                                  min="0.1"
                                  step="0.1"
                                  value={farmer.duration}
                                  onChange={(e) => updateFarmer(farmer.id, 'duration', parseFloat(e.target.value) || 0)}
                                  className="w-20 bg-gray-50 border border-gray-100 rounded px-2 py-1 text-sm font-mono focus:bg-white transition-all outline-none"
                                />
                                <span className="text-[10px] text-gray-400 font-medium">H</span>
                              </div>
                            </td>
                            <td className="px-6 py-4">
                              <div className="flex items-center gap-2">
                                <input 
                                  type="number"
                                  min="0"
                                  value={farmer.flow}
                                  onChange={(e) => updateFarmer(farmer.id, 'flow', parseFloat(e.target.value) || 0)}
                                  className="w-20 bg-gray-50 border border-gray-100 rounded px-2 py-1 text-sm font-mono focus:bg-white transition-all outline-none"
                                />
                                <span className="text-[10px] text-gray-400 font-medium">L/S</span>
                              </div>
                            </td>
                            <td className="px-6 py-4 font-mono text-sm font-semibold text-blue-600">
                              {(farmer.flow * farmer.duration * 3.6).toFixed(2)}
                            </td>
                            <td className="px-6 py-4">
                              {schedule && !isDirty ? (
                                <span className="inline-flex items-center gap-1.5 px-2.5 py-1 rounded-full text-xs font-medium bg-green-50 text-green-700">
                                  <Clock size={12} />
                                  {formatTime(schedule.startTime)}
                                </span>
                              ) : (
                                <span className="text-xs text-gray-300 italic">--:--</span>
                              )}
                            </td>
                            <td className="px-6 py-4">
                              {schedule && !isDirty ? (
                                <span className="inline-flex items-center gap-1.5 px-2.5 py-1 rounded-full text-xs font-medium bg-orange-50 text-orange-700">
                                  <ChevronRight size={12} />
                                  {formatTime(schedule.endTime)}
                                </span>
                              ) : (
                                <span className="text-xs text-gray-300 italic">--:--</span>
                              )}
                            </td>
                          </motion.tr>
                        );
                      })}
                    </AnimatePresence>
                  </tbody>
                  {farmers.length > 0 && (
                    <tfoot>
                      {isDirty && scheduledResults.length > 0 && (
                        <tr className="bg-amber-50 no-print">
                          <td colSpan={6} className="px-6 py-2 text-center">
                            <p className="text-[10px] font-bold text-amber-700 uppercase tracking-widest flex items-center justify-center gap-2">
                              <AlertCircle size={12} />
                              Modifications détectées - Relancez le calcul
                            </p>
                          </td>
                        </tr>
                      )}
                      <tr className="bg-gray-900 text-white">
                        <td colSpan={3} className="px-6 py-5 text-sm font-medium uppercase tracking-widest text-gray-400">
                          Total Volume du Tour
                        </td>
                        <td className="px-6 py-5 font-mono text-xl font-bold">
                          {totalVolume.toFixed(2)} <span className="text-sm font-light text-gray-400">m³</span>
                        </td>
                        <td colSpan={2} className="px-6 py-5 text-right">
                          <div className="inline-flex items-center gap-2 text-[10px] uppercase font-bold tracking-tighter bg-gray-800 px-3 py-1.5 rounded-lg whitespace-nowrap">
                            <Info size={12} className={scheduledResults.length > 0 && !isDirty ? "text-green-400" : "text-blue-400"} />
                            {scheduledResults.length > 0 && !isDirty ? "Planification Terminée" : "En attente"}
                          </div>
                        </td>
                      </tr>
                    </tfoot>
                  )}
                </table>
              </div>
            </div>
          </div>
        ) : (
          <div className="space-y-8">
            <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6">
              {Array.from({ length: effectiveTourDays }).map((_, dayIdx) => {
                const dayFarmers = scheduledResults.filter(f => 
                  getDayIndex(f.startTime) === dayIdx || 
                  getDayIndex(f.endTime) === dayIdx ||
                  (getDayIndex(f.startTime) < dayIdx && getDayIndex(f.endTime) > dayIdx)
                ).sort((a, b) => a.startTime - b.startTime);

                if (dayFarmers.length === 0) return null;

                const dayTotalVolume = dayFarmers.reduce((sum, f) => {
                  // Calculate volume contributed to THIS calendar day
                  // Hours from midnight of dayIdx relative to tour start (which is 9am)
                  // tourStart = 9h from midnight.
                  // Day 0 midnight = -9h relative to tour start.
                  // Day dayIdx midnight = (dayIdx * 24) - 9.
                  const dayStart = (dayIdx * 24) - 9;
                  const dayEnd = ((dayIdx + 1) * 24) - 9;
                  
                  const actualStart = Math.max(f.startTime, dayStart);
                  const actualEnd = Math.min(f.endTime, dayEnd);
                  const durationInDay = Math.max(0, actualEnd - actualStart);
                  return sum + (f.flow * durationInDay * 3.6);
                }, 0);

                return (
                  <div key={dayIdx} className="bg-white rounded-2xl border border-gray-100 shadow-sm overflow-hidden flex flex-col">
                    <div className="bg-gray-50 px-6 py-4 border-b border-gray-100 flex justify-between items-center">
                      <h3 className="font-bold text-gray-900">Jour {dayIdx + 1}</h3>
                      <span className="text-xs text-gray-400 font-mono">
                        {new Date(new Date(tourStartDate).getTime() + dayIdx * 24 * 60 * 60 * 1000).toLocaleDateString('fr-FR')}
                      </span>
                    </div>
                    <div className="p-6 space-y-4 flex-grow">
                      {dayFarmers.map(f => (
                        <div key={f.id} className="flex items-center justify-between border-b border-gray-50 pb-3 last:border-0 last:pb-0">
                          <div className="flex items-center gap-3">
                            <span className="w-8 h-8 rounded-full bg-blue-50 text-blue-600 flex items-center justify-center font-bold text-xs">
                              {f.id}
                            </span>
                            <div className="text-[11px] text-gray-600 font-mono">
                              {formatTime(f.startTime)} → {formatTime(f.endTime)}
                            </div>
                          </div>
                          <div className="text-right">
                            <div className="text-xs font-mono font-bold text-gray-900">{f.flow} L/S</div>
                            <div className="text-[10px] text-gray-400">{f.duration}h</div>
                          </div>
                        </div>
                      ))}
                    </div>
                    <div className="bg-gray-900 px-6 py-3 text-white flex justify-between items-center text-[10px] uppercase font-bold tracking-widest">
                      <span className="text-gray-400">Total</span>
                      <span>{dayTotalVolume.toFixed(1)} m³</span>
                    </div>
                  </div>
                );
              })}
            </div>
            {scheduledResults.length === 0 && (
              <div className="text-center py-20 bg-gray-50 rounded-3xl border-2 border-dashed border-gray-200">
                <p className="text-gray-400">Lancez le tour pour voir le récapitulatif.</p>
              </div>
            )}
          </div>
        )}

        {/* Footer */}
        <footer className="pt-12 pb-8 border-t border-gray-100 flex flex-col md:flex-row items-center justify-between gap-4 text-gray-400 text-sm no-print">
          <p>© 2024 IrrigaPlan. Données sauvegardées localement.</p>
          <p>
            Designed and developed with ❤️ by{" "}
            <a 
              href="https://www.facebook.com/ProfMalsahMoumen" 
              target="_blank" 
              rel="noopener noreferrer"
              className="text-blue-600 hover:text-blue-800 font-medium transition-colors"
            >
              Moumen MAlsah
            </a>
          </p>
        </footer>


      </div>
    </div>
  );
}
