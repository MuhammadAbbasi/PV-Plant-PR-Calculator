import React, { useState, useEffect } from 'react';
import {
  Chart as ChartJS,
  CategoryScale,
  LinearScale,
  PointElement,
  LineElement,
  BarElement,
  Title,
  Tooltip,
  Legend,
  Filler
} from 'chart.js';
import { Line, Bar } from 'react-chartjs-2';
import zoomPlugin from 'chartjs-plugin-zoom';
import { 
  Sun, 
  Zap, 
  Activity, 
  AlertTriangle, 
  Calendar, 
  RotateCcw, 
  CheckCircle,
  Database,
  Grid,
  TrendingUp,
  BarChart2,
  Clock,
  ChevronRight,
  ChevronLeft,
  Moon
} from 'lucide-react';

// Register ChartJS modules
ChartJS.register(
  CategoryScale,
  LinearScale,
  PointElement,
  LineElement,
  BarElement,
  Title,
  Tooltip,
  Legend,
  Filler,
  zoomPlugin
);

// Months map helper (Global)
const italianMonths = {
  "01": "Gennaio", "02": "Febbraio", "03": "Marzo", "04": "Aprile",
  "05": "Maggio", "06": "Giugno", "07": "Luglio", "08": "Agosto",
  "09": "Settembre", "10": "Ottobre", "11": "Novembre", "12": "Dicembre"
};

const monthsList = ["Gen", "Feb", "Mar", "Apr", "Mag", "Giu", "Lug", "Ago", "Set", "Ott", "Nov", "Dic"];

// Helper to get Chart.js options dynamically according to the active theme
const getChartOptions = (theme, yMin = null, yMax = null, stacked = false, zoomEnabled = false) => {
  const gridColor = theme === 'dark' ? 'rgba(255, 255, 255, 0.07)' : 'rgba(0, 0, 0, 0.08)';
  const textColor = theme === 'dark' ? '#94a3b8' : '#475569';
  const tooltipBg = theme === 'dark' ? '#181b24' : '#ffffff';
  const tooltipText = theme === 'dark' ? '#f8fafc' : '#0f172a';
  const tooltipBorder = theme === 'dark' ? 'rgba(255, 255, 255, 0.1)' : 'rgba(0, 0, 0, 0.08)';

  const options = {
    responsive: true,
    maintainAspectRatio: false,
    plugins: {
      legend: {
        labels: {
          color: textColor,
          font: { family: 'Segoe UI', size: 11, weight: 'bold' }
        }
      },
      tooltip: {
        backgroundColor: tooltipBg,
        titleColor: tooltipText,
        bodyColor: tooltipText,
        borderColor: tooltipBorder,
        borderWidth: 1,
        titleFont: { family: 'Segoe UI', weight: 'bold' },
        bodyFont: { family: 'Segoe UI' }
      }
    },
    scales: {
      x: {
        grid: { color: gridColor },
        ticks: { color: textColor, font: { family: 'Segoe UI' } },
        stacked: stacked
      },
      y: {
        grid: { color: gridColor },
        ticks: { color: textColor, font: { family: 'Segoe UI' } },
        stacked: stacked
      }
    }
  };

  if (yMin !== null) options.scales.y.min = yMin;
  if (yMax !== null) options.scales.y.max = yMax;

  if (zoomEnabled) {
    options.plugins.zoom = {
      zoom: {
        wheel: { enabled: true },
        pinch: { enabled: true },
        mode: 'x',
      },
      pan: {
        enabled: true,
        mode: 'x',
      }
    };
  }

  return options;
};

const add15Minutes = (timeStr) => {
  if (!timeStr) return '';
  const parts = timeStr.split(':');
  const h = parseInt(parts[0], 10);
  const m = parseInt(parts[1], 10);
  let newM = m + 15;
  let newH = h;
  if (newM >= 60) {
    newM -= 60;
    newH += 1;
  }
  if (newH >= 24) {
    return '24:00';
  }
  const formattedH = String(newH).padStart(2, '0');
  const formattedM = String(newM).padStart(2, '0');
  return `${formattedH}:${formattedM}`;
};

// Daily diagnostic logic for Excel status and active regulation columns
const analyzeDailyLossEvents = (intervals) => {
  if (!intervals || intervals.length === 0) return [];
  
  const events = [];
  
  // 1. Analyze Curtailment (Active Power Regulation Limit Ratio < 0.875)
  let inCurtailment = false;
  let curtailmentStart = null;
  let minRatio = 1.0;
  
  intervals.forEach((interval) => {
    const regVal = interval.active_power_regulation;
    if (regVal < 0.875) { // curtailment active (below nominal 87.6% limit)
      if (!inCurtailment) {
        inCurtailment = true;
        curtailmentStart = interval.time;
        minRatio = regVal;
      } else {
        if (regVal < minRatio) minRatio = regVal;
      }
    } else {
      if (inCurtailment) {
        events.push({
          type: 'curtailment',
          description: `Grid Curtailment (Limitazione Rete): potenza attiva limitata al ${(minRatio * 100).toFixed(1)}% dalle ${curtailmentStart.substring(0, 5)} alle ${interval.time.substring(0, 5)}`
        });
        inCurtailment = false;
      }
    }
  });
  if (inCurtailment) {
    events.push({
      type: 'curtailment',
      description: `Grid Curtailment (Limitazione Rete): potenza attiva limitata al ${(minRatio * 100).toFixed(1)}% dalle ${curtailmentStart.substring(0, 5)} alle 24:00`
    });
  }

  // 2. Analyze Inverter Outages (Irradiance > 50 W/m2 and status = 0)
  const inverterOutages = {};
  
  intervals.forEach((interval) => {
    // If grid curtailment is active in this interval, skip inverter outages (shutdown is commanded by operator)
    if (interval.active_power_regulation < 0.875) return;

    const isSunShining = (interval.irr_ref * 4000) > 50.0;
    if (!isSunShining) return; // ignore night or low irradiance intervals
    
    Object.entries(interval.inverter_statuses).forEach(([invName, statusVal]) => {
      if (statusVal === 0) { // inverter is down
        if (!inverterOutages[invName]) {
          inverterOutages[invName] = [];
        }
        inverterOutages[invName].push(interval.time);
      }
    });
  });

  // Group contiguous intervals for each inverter
  Object.entries(inverterOutages).forEach(([invName, times]) => {
    if (times.length === 0) return;
    
    // Group contiguous times
    let start = times[0];
    
    // We map times to indices of intervals to check contiguity
    const timeToIdx = {};
    intervals.forEach((interval, idx) => {
      timeToIdx[interval.time] = idx;
    });
    
    for (let i = 1; i <= times.length; i++) {
      const prevTime = times[i - 1];
      const currTime = times[i];
      const prevIdx = timeToIdx[prevTime];
      const currIdx = currTime ? timeToIdx[currTime] : -999;
      
      // If not contiguous, or end of list, output event
      if (currIdx !== prevIdx + 1) {
        events.push({
          type: 'outage',
          description: `Fermo Inverter ${invName}: spento/offline dalle ${start.substring(0, 5)} alle ${add15Minutes(prevTime)} con sole presente`
        });
        start = currTime;
      }
    }
  });

  return events;
};

export default function DashboardApp() {
  const [viewType, setViewType] = useState('month'); // 'year', 'month', 'day'
  const [theme, setTheme] = useState(() => localStorage.getItem('theme') || 'dark');
  
  // Navigation states
  const [years, setYears] = useState([]);
  const [months, setMonths] = useState([]);
  const [selectedYear, setSelectedYear] = useState('');
  const [selectedMonth, setSelectedMonth] = useState('');
  const [selectedDate, setSelectedDate] = useState(''); // YYYY-MM-DD
  
  // Data states
  const [syncStatus, setSyncStatus] = useState(null);
  const [yearlyData, setYearlyData] = useState([]);
  const [monthlyData, setMonthlyData] = useState([]);
  const [dailyData, setDailyData] = useState(null);
  
  // Loading & error states
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState(null);

  // Sync class theme on mount & change
  useEffect(() => {
    if (theme === 'light') {
      document.body.classList.add('light-theme');
    } else {
      document.body.classList.remove('light-theme');
    }
    localStorage.setItem('theme', theme);
  }, [theme]);

  // Fetch initial configurations & sync status
  useEffect(() => {
    fetchSyncStatus();
    fetchYears();
    
    // Poll sync status if indexing is active
    const interval = setInterval(() => {
      fetchSyncStatus();
    }, 5000);
    return () => clearInterval(interval);
  }, []);

  const fetchSyncStatus = async () => {
    try {
      const res = await fetch('/api/sync-status');
      const data = await res.json();
      setSyncStatus(data);
    } catch (err) {
      console.error("Error fetching sync status:", err);
    }
  };

  const triggerSync = async () => {
    try {
      await fetch('/api/sync', { method: 'POST' });
      fetchSyncStatus();
    } catch (err) {
      console.error("Error triggering sync:", err);
    }
  };

  const fetchYears = async () => {
    try {
      const res = await fetch('/api/years');
      const data = await res.json();
      setYears(data);
      if (data.length > 0) {
        // Default to latest year
        setSelectedYear(data[0]);
      }
    } catch (err) {
      console.error("Error fetching years:", err);
    }
  };

  // Fetch months when selectedYear changes
  useEffect(() => {
    if (!selectedYear) return;
    const fetchMonths = async () => {
      try {
        const res = await fetch(`/api/months?year=${selectedYear}`);
        const data = await res.json();
        setMonths(data);
        if (data.length > 0) {
          // Default to latest month in list
          setSelectedMonth(data[data.length - 1]);
        }
      } catch (err) {
        console.error("Error fetching months:", err);
      }
    };
    fetchMonths();
  }, [selectedYear]);

  // Fetch data depending on active view & selections
  useEffect(() => {
    if (viewType === 'year' && selectedYear) {
      fetchYearlyData();
    } else if (viewType === 'month' && selectedYear && selectedMonth) {
      fetchMonthlyData();
    } else if (viewType === 'day' && selectedDate) {
      fetchDailyData();
    }
  }, [viewType, selectedYear, selectedMonth, selectedDate]);

  const fetchYearlyData = async () => {
    setLoading(true);
    setError(null);
    try {
      const res = await fetch(`/api/yearly-summary?year=${selectedYear}`);
      if (!res.ok) throw new Error("Errore nel caricamento dei dati annuali");
      const data = await res.json();
      setYearlyData(data);
    } catch (err) {
      setError(err.message);
    } finally {
      setLoading(false);
    }
  };

  const fetchMonthlyData = async () => {
    setLoading(true);
    setError(null);
    try {
      const res = await fetch(`/api/monthly-data?year=${selectedYear}&month=${selectedMonth}`);
      if (!res.ok) throw new Error("Errore nel caricamento dei dati mensili");
      const data = await res.json();
      setMonthlyData(data);
      if (data.length > 0) {
        setSelectedDate(data[0].date);
      }
    } catch (err) {
      setError(err.message);
    } finally {
      setLoading(false);
    }
  };

  const fetchDailyData = async () => {
    setLoading(true);
    setError(null);
    try {
      const res = await fetch(`/api/daily-data?date=${selectedDate}`);
      if (!res.ok) throw new Error("Errore nel caricamento del file giornaliero");
      const data = await res.json();
      setDailyData(data);
    } catch (err) {
      setError(err.message);
    } finally {
      setLoading(false);
    }
  };

  const navigateToDay = (date) => {
    setSelectedDate(date);
    setViewType('day');
  };

  // Navigations arrow logic
  const navigateYear = (direction) => {
    if (!selectedYear || years.length === 0) return;
    const idx = years.indexOf(selectedYear);
    if (idx === -1) return;
    
    // Note: years are sorted descending e.g. ["2026", "2025"]
    // direction -1 = backward in time (add to index)
    // direction 1 = forward in time (subtract from index)
    const nextIdx = idx - direction;
    if (nextIdx >= 0 && nextIdx < years.length) {
      setSelectedYear(years[nextIdx]);
    }
  };

  const navigateMonth = (direction) => {
    if (!selectedYear || !selectedMonth) return;
    let y = parseInt(selectedYear);
    let m = parseInt(selectedMonth);
    
    m += direction;
    if (m > 12) {
      m = 1;
      y += 1;
    } else if (m < 1) {
      m = 12;
      y -= 1;
    }
    
    const newYearStr = String(y);
    const newMonthStr = String(m).padStart(2, '0');
    
    if (years.includes(newYearStr)) {
      setSelectedYear(newYearStr);
      // Wait for useEffect of selectedYear to fetch available months
      // Directly override selectedMonth state if months list updates or is already available
      setSelectedMonth(newMonthStr);
    }
  };

  const navigateDay = (direction) => {
    if (!selectedDate) return;
    const current = new Date(selectedDate);
    if (isNaN(current.getTime())) return;
    
    current.setDate(current.getDate() + direction);
    const yyyy = current.getFullYear();
    const mm = String(current.getMonth() + 1).padStart(2, '0');
    const dd = String(current.getDate()).padStart(2, '0');
    setSelectedDate(`${yyyy}-${mm}-${dd}`);
  };

  // Aggregations
  const monthlyAverages = React.useMemo(() => {
    if (monthlyData.length === 0) return { pr: 0, scada: 0, target: 83.2, energy: 0, loss: 0 };
    let sumComp = 0;
    let sumScada = 0;
    let sumTarget = 0;
    let totalEnergy = 0;
    let totalLoss = 0;
    let daysWithTarget = 0;
    
    monthlyData.forEach(d => {
      sumComp += d.pr_compensated;
      sumScada += d.pr_scada;
      totalEnergy += d.energy;
      totalLoss += (d.loss_tx1 + d.loss_tx2 + d.loss_tx3);
      if (d.pvsyst_pr_target) {
        sumTarget += d.pvsyst_pr_target;
        daysWithTarget++;
      }
    });
    
    return {
      pr: sumComp / monthlyData.length,
      scada: sumScada / monthlyData.length,
      target: daysWithTarget > 0 ? (sumTarget / daysWithTarget) : 83.2,
      energy: totalEnergy,
      loss: totalLoss
    };
  }, [monthlyData]);

  const yearlyTotals = React.useMemo(() => {
    if (yearlyData.length === 0) return { energy: 0, pr: 0, loss: 0, loss_tx1: 0, loss_tx2: 0, loss_tx3: 0 };
    let totalEnergy = 0;
    let sumPr = 0;
    let totalLoss = 0;
    let totalLossTx1 = 0;
    let totalLossTx2 = 0;
    let totalLossTx3 = 0;
    
    yearlyData.forEach(m => {
      totalEnergy += m.total_energy;
      sumPr += m.avg_pr_compensated;
      totalLoss += (m.total_loss_tx1 + m.total_loss_tx2 + m.total_loss_tx3);
      totalLossTx1 += m.total_loss_tx1;
      totalLossTx2 += m.total_loss_tx2;
      totalLossTx3 += m.total_loss_tx3;
    });
    return {
      energy: totalEnergy,
      pr: sumPr / yearlyData.length,
      loss: totalLoss,
      loss_tx1: totalLossTx1,
      loss_tx2: totalLossTx2,
      loss_tx3: totalLossTx3
    };
  }, [yearlyData]);

  return (
    <div className="app-container">
      {/* Header */}
      <header className="app-header">
        <div className="logo-section">
          <Sun className="logo-icon" size={32} />
          <div>
            <h1 className="logo-title">Mazara 01</h1>
            <span className="logo-sub">GET S.r.l. • PR Dashboard</span>
          </div>
        </div>

        {/* Navigation Selector */}
        <div className="nav-controls">
          <div className="btn-group">
            <button 
              className={`btn-tab ${viewType === 'year' ? 'active' : ''}`}
              onClick={() => setViewType('year')}
            >
              Anno
            </button>
            <button 
              className={`btn-tab ${viewType === 'month' ? 'active' : ''}`}
              onClick={() => setViewType('month')}
            >
              Mese
            </button>
            <button 
              className={`btn-tab ${viewType === 'day' ? 'active' : ''}`}
              onClick={() => setViewType('day')}
            >
              Giorno
            </button>
          </div>

          {/* Temporal filters with navigation arrows */}
          {viewType === 'year' && (
            <div style={{ display: 'flex', alignItems: 'center', gap: '0.25rem' }}>
              <button 
                className="btn" 
                style={{ padding: '0.4rem' }} 
                onClick={() => navigateYear(-1)} 
                title="Anno precedente"
              >
                <ChevronLeft size={16} />
              </button>
              <div className="select-wrapper">
                <select 
                  value={selectedYear} 
                  onChange={(e) => setSelectedYear(e.target.value)}
                  className="nav-select"
                >
                  {years.map(y => <option key={y} value={y}>{y}</option>)}
                </select>
              </div>
              <button 
                className="btn" 
                style={{ padding: '0.4rem' }} 
                onClick={() => navigateYear(1)} 
                title="Anno successivo"
              >
                <ChevronRight size={16} />
              </button>
            </div>
          )}

          {viewType === 'month' && (
            <div style={{ display: 'flex', alignItems: 'center', gap: '0.25rem' }}>
              <button 
                className="btn" 
                style={{ padding: '0.4rem' }} 
                onClick={() => navigateMonth(-1)} 
                title="Mese precedente"
              >
                <ChevronLeft size={16} />
              </button>
              <div className="select-wrapper" style={{ marginRight: '0.25rem' }}>
                <select 
                  value={selectedYear} 
                  onChange={(e) => setSelectedYear(e.target.value)}
                  className="nav-select"
                >
                  {years.map(y => <option key={y} value={y}>{y}</option>)}
                </select>
              </div>
              <div className="select-wrapper">
                <select 
                  value={selectedMonth} 
                  onChange={(e) => setSelectedMonth(e.target.value)}
                  className="nav-select"
                >
                  {months.map(m => (
                    <option key={m} value={m}>{italianMonths[m] || m}</option>
                  ))}
                </select>
              </div>
              <button 
                className="btn" 
                style={{ padding: '0.4rem' }} 
                onClick={() => navigateMonth(1)} 
                title="Mese successivo"
              >
                <ChevronRight size={16} />
              </button>
            </div>
          )}

          {viewType === 'day' && (
            <div style={{ display: 'flex', alignItems: 'center', gap: '0.25rem' }}>
              <button 
                className="btn" 
                style={{ padding: '0.4rem' }} 
                onClick={() => navigateDay(-1)} 
                title="Giorno precedente"
              >
                <ChevronLeft size={16} />
              </button>
              <input 
                type="date"
                value={selectedDate}
                onChange={(e) => setSelectedDate(e.target.value)}
                className="nav-select"
                style={{ paddingRight: '1rem' }}
              />
              <button 
                className="btn" 
                style={{ padding: '0.4rem' }} 
                onClick={() => navigateDay(1)} 
                title="Giorno successivo"
              >
                <ChevronRight size={16} />
              </button>
            </div>
          )}

          {/* Theme Toggle Button */}
          <button 
            className="btn"
            onClick={() => setTheme(theme === 'dark' ? 'light' : 'dark')}
            title={theme === 'dark' ? 'Attiva tema chiaro' : 'Attiva tema scuro'}
            style={{ display: 'inline-flex', alignItems: 'center', gap: '0.4rem' }}
          >
            {theme === 'dark' ? <Sun size={14} /> : <Moon size={14} />}
            {theme === 'dark' ? 'Chiaro' : 'Scuro'}
          </button>

          {/* Manual cache database trigger */}
          <button 
            className="btn"
            onClick={triggerSync}
            disabled={syncStatus?.is_syncing}
            title="Aggiorna i dati scansionando i file Excel modificati"
          >
            <RotateCcw size={16} />
            Sincronizza
          </button>
        </div>
      </header>

      <main className="app-content">
        {/* Sync/indexing progress notification banner */}
        {syncStatus?.is_syncing && (
          <div className="sync-banner">
            <div style={{ display: 'flex', alignItems: 'center', gap: '0.75rem' }}>
              <Database className="logo-icon" size={18} />
              <span>
                <strong>Scansione cache database attiva:</strong> Sincronizzazione dei file Excel in corso... ({syncStatus.processed} di {syncStatus.total} file)
              </span>
            </div>
            <span style={{ fontSize: '0.8rem', color: 'var(--text-secondary)', fontStyle: 'italic' }}>
              Elaborazione di: {syncStatus.current_file}
            </span>
          </div>
        )}

        {/* Global Loading Spinner */}
        {loading ? (
          <div className="loader-container">
            <div className="spinner"></div>
            <p style={{ color: 'var(--text-secondary)' }}>Caricamento in corso...</p>
          </div>
        ) : error ? (
          <div className="card" style={{ borderColor: 'var(--danger)', padding: '2rem', alignItems: 'center', gap: '1rem' }}>
            <AlertTriangle color="var(--danger)" size={48} />
            <h3 style={{ fontSize: '1.25rem' }}>Errore di Caricamento Dati</h3>
            <p style={{ color: 'var(--text-secondary)', textAlign: 'center' }}>{error}</p>
            <button className="btn btn-primary" onClick={() => {
              if (viewType === 'year') fetchYearlyData();
              else if (viewType === 'month') fetchMonthlyData();
              else fetchDailyData();
            }}>
              Riprova
            </button>
          </div>
        ) : (
          <>
            {viewType === 'year' && renderYearlyView(yearlyData, yearlyTotals, selectedYear, theme)}
            {viewType === 'month' && renderMonthlyView(monthlyData, monthlyAverages, selectedYear, selectedMonth, navigateToDay, theme)}
            {viewType === 'day' && dailyData && renderDailyView(dailyData, theme)}
          </>
        )}
      </main>
    </div>
  );
}

// ----------------------------------------------------
// YEARLY VIEW RENDERER
// ----------------------------------------------------
function renderYearlyView(data, totals, year, theme) {
  // Line Chart: Monthly PR trends
  const prChartData = {
    labels: data.map(m => monthsList[parseInt(m.month) - 1]),
    datasets: [
      {
        label: 'Compensated PR [%]',
        data: data.map(m => m.avg_pr_compensated),
        borderColor: '#ea580c',
        backgroundColor: 'rgba(234, 88, 12, 0.1)',
        tension: 0.1,
        fill: true,
        pointRadius: 4
      },
      {
        label: 'SCADA PR [%]',
        data: data.map(m => m.avg_pr_scada),
        borderColor: '#94a3b8',
        borderDash: [5, 5],
        tension: 0.1,
        pointRadius: 4
      }
    ]
  };

  // Line Chart: Monthly Actual Energy vs Potential Energy (shaded difference)
  const energyChartData = {
    labels: data.map(m => monthsList[parseInt(m.month) - 1]),
    datasets: [
      {
        label: 'Energia Effettiva Prodotta [kWh]',
        data: data.map(m => m.total_energy),
        borderColor: '#10b981',
        backgroundColor: 'rgba(16, 185, 129, 0.05)',
        tension: 0.15,
        fill: true,
        pointRadius: 4
      },
      {
        label: 'Energia Potenziale [kWh]',
        data: data.map(m => m.total_energy + m.total_loss_tx1 + m.total_loss_tx2 + m.total_loss_tx3),
        borderColor: '#ea580c',
        backgroundColor: 'rgba(234, 88, 12, 0.25)', // Shading orange
        tension: 0.15,
        fill: 0, // Shading difference
        pointRadius: 4
      }
    ]
  };

  return (
    <>
      {/* Yearly KPIs */}
      <div className="kpi-grid">
        <div className="kpi-card success">
          <div className="kpi-header">
            <span>Energia Prodotta Totale</span>
            <Zap size={18} color="var(--success)" />
          </div>
          <div className="kpi-value">{Math.round(totals.energy).toLocaleString('it-IT')} kWh</div>
          <div className="kpi-subtitle">Anno Solare {year}</div>
        </div>

        <div className="kpi-card">
          <div className="kpi-header">
            <span>Compensated PR Medio</span>
            <TrendingUp size={18} color="var(--accent)" />
          </div>
          <div className="kpi-value">{totals.pr.toFixed(2)} %</div>
          <div className="kpi-subtitle">Media aritmetica dei mesi</div>
        </div>

        <div className="kpi-card danger">
          <div className="kpi-header">
            <span>Perdite di Energia Totali</span>
            <AlertTriangle size={18} color="var(--danger)" />
          </div>
          <div className="kpi-value">{Math.round(totals.loss).toLocaleString('it-IT')} kWh</div>
          <div className="kpi-subtitle">Somma delle perdite TX1, TX2, TX3</div>
        </div>
      </div>

      {/* Yearly Diagrams */}
      <div className="dashboard-row">
        <div className="card">
          <h2 className="card-title">Andamento PR Mensile</h2>
          <div className="chart-container">
            <Line 
              data={prChartData} 
              options={getChartOptions(theme)} 
            />
          </div>
        </div>

        <div className="card">
          <h2 className="card-title">
            <span>Analisi Energia Potenziale vs Effettiva (Perdite Shaded)</span>
            <span style={{ fontSize: '0.8rem', color: 'var(--warning)', fontWeight: 'bold' }}>Area Arancione = Energia Persa</span>
          </h2>
          <div className="chart-container">
            <Line 
              data={energyChartData} 
              options={getChartOptions(theme)} 
            />
          </div>
        </div>
      </div>

      {/* Yearly Details Table */}
      <div className="card">
        <h2 className="card-title">Tabella Dettaglio Mensile - Andamento Annuale {year}</h2>
        <div className="table-container">
          <table className="data-table">
            <thead>
              <tr>
                <th>Mese</th>
                <th>Energia Prodotta [kWh]</th>
                <th>Irradiance Ref [kWh/m²]</th>
                <th>Compensated PR Medio [%]</th>
                <th>SCADA PR Medio [%]</th>
                <th>Raw PR Medio [%]</th>
                <th>Perdite TX1 [kWh]</th>
                <th>Perdite TX2 [kWh]</th>
                <th>Perdite TX3 [kWh]</th>
                <th>Perdite Totali [kWh]</th>
              </tr>
            </thead>
            <tbody>
              {data.map((m) => {
                const totalLoss = m.total_loss_tx1 + m.total_loss_tx2 + m.total_loss_tx3;
                return (
                  <tr key={m.month}>
                    <td><strong>{italianMonths[m.month] || m.month}</strong></td>
                    <td>{Math.round(m.total_energy).toLocaleString('it-IT')}</td>
                    <td>{m.total_irradiance.toFixed(3)}</td>
                    <td><strong style={{ color: 'var(--accent)' }}>{m.avg_pr_compensated.toFixed(2)} %</strong></td>
                    <td>{m.avg_pr_scada.toFixed(2)} %</td>
                    <td>{m.avg_pr_raw.toFixed(2)} %</td>
                    <td>{Math.round(m.total_loss_tx1).toLocaleString('it-IT')}</td>
                    <td>{Math.round(m.total_loss_tx2).toLocaleString('it-IT')}</td>
                    <td>{Math.round(m.total_loss_tx3).toLocaleString('it-IT')}</td>
                    <td><strong style={{ color: 'var(--danger)' }}>{Math.round(totalLoss).toLocaleString('it-IT')}</strong></td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      </div>

      {/* Diagnostica Perdite Annue */}
      <div className="card">
        <h2 className="card-title">Dettaglio Diagnostico delle Perdite Energetiche Annue</h2>
        <div className="dashboard-row" style={{ gridTemplateColumns: '1fr 1fr' }}>
          <div>
            <h3 style={{ fontSize: '1rem', marginBottom: '0.5rem', color: 'var(--text-secondary)' }}>Ripartizione delle perdite per Trasformatore:</h3>
            <ul style={{ paddingLeft: '1.25rem', lineHeight: '1.75' }}>
              <li><strong>Trasformatore TX1 (Inverter O-Z)</strong>: {Math.round(totals.loss_tx1).toLocaleString('it-IT')} kWh persi.</li>
              <li><strong>Trasformatore TX2 (Inverter AB-AM)</strong>: {Math.round(totals.loss_tx2).toLocaleString('it-IT')} kWh persi.</li>
              <li><strong>Trasformatore TX3 (Inverter AO-AZ)</strong>: {Math.round(totals.loss_tx3).toLocaleString('it-IT')} kWh persi.</li>
            </ul>
          </div>
          <div>
            <h3 style={{ fontSize: '1rem', marginBottom: '0.5rem', color: 'var(--text-secondary)' }}>Analisi dei mesi critici:</h3>
            <p style={{ lineHeight: '1.5', color: 'var(--text-secondary)' }}>
              Il mese con la maggior perdita energetica nell'anno {year} è stato <strong>{
                (() => {
                  if (data.length === 0) return 'N/D';
                  const maxLossMonth = [...data].sort((a, b) => (b.total_loss_tx1+b.total_loss_tx2+b.total_loss_tx3) - (a.total_loss_tx1+a.total_loss_tx2+a.total_loss_tx3))[0];
                  const lossVal = maxLossMonth.total_loss_tx1 + maxLossMonth.total_loss_tx2 + maxLossMonth.total_loss_tx3;
                  return `${italianMonths[maxLossMonth.month]} (con ${Math.round(lossVal).toLocaleString('it-IT')} kWh persi)`;
                })()
              }</strong>.
              Si consiglia di verificare lo stato di manutenzione degli inverter afferenti a quel trasformatore e i registri di limitazione del carico di rete per quel periodo.
            </p>
          </div>
        </div>
      </div>
    </>
  );
}

// ----------------------------------------------------
// MONTHLY VIEW RENDERER
// ----------------------------------------------------
function renderMonthlyView(data, avgs, year, month, onDaySelect, theme) {
  if (data.length === 0) {
    return (
      <div className="card" style={{ padding: '3rem', textAlign: 'center', color: 'var(--text-secondary)' }}>
        Nessun dato registrato per questo mese.
      </div>
    );
  }

  // Plot Daily PR values with Average overlay & Benchmark lines
  const dailyLabels = data.map(d => {
    const parts = d.date.split('-');
    return parts[2]; // just day number
  });

  const prData = {
    labels: dailyLabels,
    datasets: [
      {
        label: 'Daily Compensated PR [%]',
        data: data.map(d => d.pr_compensated),
        borderColor: '#ea580c',
        backgroundColor: '#ea580c',
        pointRadius: 4,
        tension: 0.1,
        order: 1
      },
      {
        label: 'PVSyst PR Target (Storico)',
        data: data.map(d => d.pvsyst_pr_target || avgs.target),
        borderColor: '#f59e0b',
        borderDash: [8, 4],
        pointRadius: 0,
        tension: 0,
        fill: false,
        borderWidth: 2,
        order: 3
      },
      {
        label: 'Media Mensile Compensata',
        data: Array(data.length).fill(avgs.pr),
        borderColor: 'rgba(16, 185, 129, 0.7)',
        borderWidth: 1.5,
        pointRadius: 0,
        tension: 0,
        order: 2
      }
    ]
  };

  // Line Chart: Daily Actual Energy vs Potential Energy (shaded orange)
  const energyChartData = {
    labels: dailyLabels,
    datasets: [
      {
        label: 'Energia Effettiva Prodotta [kWh]',
        data: data.map(d => d.energy),
        borderColor: '#10b981',
        backgroundColor: 'rgba(16, 185, 129, 0.05)',
        tension: 0.1,
        fill: true,
        pointRadius: 3
      },
      {
        label: 'Energia Potenziale [kWh]',
        data: data.map(d => d.energy + d.loss_tx1 + d.loss_tx2 + d.loss_tx3),
        borderColor: '#ea580c',
        backgroundColor: 'rgba(234, 88, 12, 0.25)', // orange shading
        tension: 0.1,
        fill: 0, // fill to dataset 0
        pointRadius: 3
      }
    ]
  };

  return (
    <>
      {/* Monthly KPIs */}
      <div className="kpi-grid">
        <div className="kpi-card">
          <div className="kpi-header">
            <span>Compensated PR Medio</span>
            <TrendingUp size={18} color="var(--accent)" />
          </div>
          <div className="kpi-value">{avgs.pr.toFixed(2)} %</div>
          <div className="kpi-subtitle">Media mensile compensata</div>
        </div>

        <div className="kpi-card success">
          <div className="kpi-header">
            <span>Energia Prodotta Totale</span>
            <Zap size={18} color="var(--success)" />
          </div>
          <div className="kpi-value">{(avgs.energy).toLocaleString('it-IT')} kWh</div>
          <div className="kpi-subtitle">Produzione attiva del mese</div>
        </div>

        <div className="kpi-card warning">
          <div className="kpi-header">
            <span>Target Storico PVSyst</span>
            <CheckCircle size={18} color="var(--warning)" />
          </div>
          <div className="kpi-value">{avgs.target.toFixed(2)} %</div>
          <div className="kpi-subtitle">Benchmark storico di riferimento</div>
        </div>

        <div className="kpi-card danger">
          <div className="kpi-header">
            <span>Perdite di Energia Totali</span>
            <AlertTriangle size={18} color="var(--danger)" />
          </div>
          <div className="kpi-value">{(avgs.loss).toLocaleString('it-IT')} kWh</div>
          <div className="kpi-subtitle">Perdite per downtime e limitazioni</div>
        </div>
      </div>

      {/* Monthly Diagrams */}
      <div className="dashboard-row">
        <div className="card">
          <h2 className="card-title">Compensated PR Giornaliero vs Target</h2>
          <div className="chart-container">
            <Line 
              data={prData} 
              options={getChartOptions(theme, null, null, false, true)} 
            />
          </div>
        </div>

        <div className="card">
          <h2 className="card-title">
            <span>Analisi Energia Potenziale vs Effettiva (Perdite Shaded)</span>
            <span style={{ fontSize: '0.8rem', color: 'var(--warning)', fontWeight: 'bold' }}>Area Arancione = Energia Persa</span>
          </h2>
          <div className="chart-container">
            <Line 
              data={energyChartData} 
              options={getChartOptions(theme, null, null, false, true)} 
            />
          </div>
        </div>
      </div>

      {/* Side-by-side Monthly Comparison View */}
      <div className="card">
        <div className="card-title">
          <span>Tabella Dettaglio Giornaliero - Confronto Trend Mensile</span>
          <span style={{ fontSize: '0.8rem', color: 'var(--text-secondary)' }}>Clicca su una riga per visualizzare il grafico orario</span>
        </div>
        <div className="table-container">
          <table className="data-table">
            <thead>
              <tr>
                <th>Data</th>
                <th>Irradiance TX1 [kWh/m²]</th>
                <th>Irradiance TX3 [kWh/m²]</th>
                <th>Reference POA [kWh/m²]</th>
                <th>Energia Attiva [kWh]</th>
                <th>Uncompensated PR [%]</th>
                <th>Compensated PR [%]</th>
                <th>Disponibilità Esterna [%]</th>
                <th>Perdite TX1 [kWh]</th>
                <th>Perdite TX2 [kWh]</th>
                <th>Perdite TX3 [kWh]</th>
                <th>Azioni</th>
              </tr>
            </thead>
            <tbody>
              {data.map((d) => {
                const parts = d.date.split('-');
                const formattedDate = `${parts[2]}/${parts[1]}/${parts[0]}`;
                
                let prColor = '';
                if (d.pr_compensated >= d.pvsyst_pr_target) prColor = 'var(--success)';
                else if (d.pr_compensated < 60) prColor = 'var(--danger)';
                else prColor = 'var(--warning)';

                return (
                  <tr key={d.date} style={{ cursor: 'pointer' }} onClick={() => onDaySelect(d.date)}>
                    <td><strong>{formattedDate}</strong></td>
                    <td>{d.irradiance_tx1.toFixed(3)}</td>
                    <td>{d.irradiance_tx3.toFixed(3)}</td>
                    <td>{d.irradiance_ref.toFixed(3)}</td>
                    <td>{d.energy.toLocaleString('it-IT')}</td>
                    <td>{d.pr_scada.toFixed(2)} %</td>
                    <td>
                      <span style={{ color: prColor, fontWeight: '700' }}>
                        {d.pr_compensated.toFixed(2)} %
                      </span>
                    </td>
                    <td>
                      <span className={`badge ${d.availability >= 95 ? 'badge-success' : d.availability >= 80 ? 'badge-warning' : 'badge-danger'}`}>
                        {d.availability.toFixed(1)} %
                      </span>
                    </td>
                    <td>{d.loss_tx1.toFixed(1)}</td>
                    <td>{d.loss_tx2.toFixed(1)}</td>
                    <td>{d.loss_tx3.toFixed(1)}</td>
                    <td style={{ color: 'var(--accent)' }}>
                      <span style={{ display: 'flex', alignItems: 'center', gap: '2px', fontSize: '0.8rem', fontWeight: 600 }}>
                        Grafico <ChevronRight size={14} />
                      </span>
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      </div>

      {/* Monthly Diagnostica */}
      <div className="card">
        <h2 className="card-title">Dettaglio Diagnostico delle Perdite Energetiche Mensili</h2>
        <div className="dashboard-row" style={{ gridTemplateColumns: '1fr 1fr' }}>
          <div>
            <h3 style={{ fontSize: '1rem', marginBottom: '0.5rem', color: 'var(--text-secondary)' }}>Ripartizione perdite del mese:</h3>
            <ul style={{ paddingLeft: '1.25rem', lineHeight: '1.75' }}>
              <li><strong>Trasformatore TX1 (Inverter campi O-Z)</strong>: {Math.round(data.reduce((a, b) => a + b.loss_tx1, 0)).toLocaleString('it-IT')} kWh</li>
              <li><strong>Trasformatore TX2 (Inverter campi AB-AM)</strong>: {Math.round(data.reduce((a, b) => a + b.loss_tx2, 0)).toLocaleString('it-IT')} kWh</li>
              <li><strong>Trasformatore TX3 (Inverter campi AO-AZ)</strong>: {Math.round(data.reduce((a, b) => a + b.loss_tx3, 0)).toLocaleString('it-IT')} kWh</li>
            </ul>
          </div>
          <div>
            <h3 style={{ fontSize: '1rem', marginBottom: '0.5rem', color: 'var(--text-secondary)' }}>Giorni con perdite critiche:</h3>
            <ul style={{ paddingLeft: '1.25rem', lineHeight: '1.5' }}>
              {[...data]
                .sort((a, b) => (b.loss_tx1+b.loss_tx2+b.loss_tx3) - (a.loss_tx1+a.loss_tx2+a.loss_tx3))
                .slice(0, 3)
                .map((day) => {
                  const parts = day.date.split('-');
                  const dateStr = `${parts[2]}/${parts[1]}/${parts[0]}`;
                  const dayLoss = day.loss_tx1 + day.loss_tx2 + day.loss_tx3;
                  return (
                    <li key={day.date} style={{ marginBottom: '0.25rem' }}>
                      <strong>{dateStr}</strong>: {Math.round(dayLoss).toLocaleString('it-IT')} kWh persi (
                      {day.loss_tx1 > day.loss_tx2 && day.loss_tx1 > day.loss_tx3 ? 'TX1 prevalente' :
                       day.loss_tx2 > day.loss_tx1 && day.loss_tx2 > day.loss_tx3 ? 'TX2 prevalente' : 'TX3 prevalente'}
                      )
                    </li>
                  );
                })}
            </ul>
          </div>
        </div>
      </div>
    </>
  );
}

// ----------------------------------------------------
// DAILY VIEW RENDERER (15-min Intervals)
// ----------------------------------------------------
function renderDailyView(daily, theme) {
  const summary = daily.summary;
  const intervals = daily.intervals;
  
  const parts = summary.date.split('-');
  const formattedTitleDate = `${parts[2]}/${parts[1]}/${parts[0]}`;

  // Line Chart: 15-minute Irradiance comparison
  const irrLineData = {
    labels: intervals.map(i => i.time.substring(0, 5)),
    datasets: [
      {
        label: 'Irradiance TX1 [W/m²]',
        data: intervals.map(i => i.irr_tx1_w),
        borderColor: 'rgba(56, 189, 248, 0.65)',
        borderWidth: 1.5,
        pointRadius: 0,
        tension: 0.1
      },
      {
        label: 'Irradiance TX3 [W/m²]',
        data: intervals.map(i => i.irr_tx3_w),
        borderColor: 'rgba(148, 163, 184, 0.65)',
        borderWidth: 1.5,
        pointRadius: 0,
        tension: 0.1
      },
      {
        label: 'Reference POA [W/m²]',
        data: intervals.map(i => i.irr_ref * 4000), 
        borderColor: '#ea580c',
        borderWidth: 2,
        pointRadius: 0,
        tension: 0.1,
        fill: false
      }
    ]
  };

  // Line Chart: 15-minute Actual vs Potential Active Power (shaded difference in kW)
  const energyChartData = {
    labels: intervals.map(i => i.time.substring(0, 5)),
    datasets: [
      {
        label: 'Potenza Effettiva Prodotta [kW]',
        data: intervals.map(i => i.energy * 4),
        borderColor: '#10b981',
        backgroundColor: 'rgba(16, 185, 129, 0.05)',
        tension: 0.1,
        fill: true,
        pointRadius: 0
      },
      {
        label: 'Potenza Potenziale [kW]',
        data: intervals.map(i => (i.energy + i.loss_tx1 + i.loss_tx2 + i.loss_tx3) * 4),
        borderColor: '#ea580c',
        backgroundColor: 'rgba(234, 88, 12, 0.25)', // Shading orange
        tension: 0.1,
        fill: 0, // Shading difference
        pointRadius: 0
      }
    ]
  };

  const noonInterval = intervals.find(i => i.time.startsWith('12:00:')) || intervals[Math.floor(intervals.length / 2)];
  const isCurtailedAtNoon = noonInterval && noonInterval.active_power_regulation < 0.875;
  const statusValues = noonInterval ? Object.entries(noonInterval.inverter_statuses) : [];

  const lossEvents = analyzeDailyLossEvents(intervals);

  return (
    <>
      {/* Daily KPIs */}
      <div className="kpi-grid">
        <div className="kpi-card success">
          <div className="kpi-header">
            <span>Compensated PR</span>
            <CheckCircle size={18} color="var(--success)" />
          </div>
          <div className="kpi-value">{summary.compensated_pr.toFixed(2)} %</div>
          <div className="kpi-subtitle">Generato dal file del {formattedTitleDate}</div>
        </div>

        <div className="kpi-card">
          <div className="kpi-header">
            <span>PR Uncompensated (SCADA)</span>
            <TrendingUp size={18} color="var(--accent)" />
          </div>
          <div className="kpi-value">{summary.uncompensated_pr.toFixed(2)} %</div>
          <div className="kpi-subtitle">Performance grezza senza compensazioni</div>
        </div>

        <div className="kpi-card warning">
          <div className="kpi-header">
            <span>PVSyst Target</span>
            <Zap size={18} color="var(--warning)" />
          </div>
          <div className="kpi-value">{(summary.pvsyst_pr_target || 83.2).toFixed(2)} %</div>
          <div className="kpi-subtitle">Valore di riferimento programmato</div>
        </div>

        <div className="kpi-card">
          <div className="kpi-header">
            <span>Valori Analizzati</span>
            <Clock size={18} color="var(--text-secondary)" />
          </div>
          <div className="kpi-value">{summary.valid_poa_values} / {summary.total_values}</div>
          <div className="kpi-subtitle">Intervalli con Irradiance &gt; {summary.min_irr_threshold} W/m²</div>
        </div>
      </div>

      {/* Hourly diagrams */}
      <div className="dashboard-row">
        <div className="card">
          <h2 className="card-title">Confronto Irradiance 15 minuti - {formattedTitleDate}</h2>
          <div className="chart-container">
            <Line 
              data={irrLineData} 
              options={getChartOptions(theme, null, null, false, true)} 
            />
          </div>
        </div>

        <div className="card">
          <h2 className="card-title">
            <span>Analisi Potenza Potenziale vs Effettiva - {formattedTitleDate}</span>
            <span style={{ fontSize: '0.8rem', color: 'var(--warning)', fontWeight: 'bold' }}>Area Arancione = Potenza Persa [kW]</span>
          </h2>
          <div className="chart-container">
            <Line 
              data={energyChartData} 
              options={getChartOptions(theme)} 
            />
          </div>
        </div>
      </div>

      {/* Status Matrix Grid at Noon */}
      <div className="card">
        <h2 className="card-title">
          <span>Stato di Funzionamento Inverter (Ore 12:00 del {formattedTitleDate})</span>
          <span style={{ fontSize: '0.8rem', color: 'var(--text-secondary)' }}>Valori di status: 0 = offline/curtailed, &gt;0 = attivo/parziale</span>
        </h2>
        
        <div className="inverters-status-grid">
          {statusValues.map(([invName, statusVal]) => {
            let tileBg = 'var(--bg-secondary)';
            let tileBorder = 'var(--border-color)';
            let textColor = 'var(--text-primary)';
            let label = statusVal > 0 ? `${statusVal.toFixed(1)}` : 'DOWNTIME';
            
            if (statusVal === 0) {
              if (isCurtailedAtNoon) {
                tileBg = 'var(--warning-bg)';
                tileBorder = 'var(--warning)';
                textColor = 'var(--warning)';
                label = 'LIMITATO';
              } else {
                tileBg = 'var(--danger-bg)';
                tileBorder = 'var(--danger)';
                textColor = 'var(--danger)';
                label = 'DOWNTIME';
              }
            } else if (statusVal > 0) {
              tileBg = 'var(--success-bg)';
              tileBorder = 'var(--success)';
              textColor = 'var(--success)';
            }
            
            return (
              <div 
                key={invName} 
                className="inverter-tile" 
                style={{ backgroundColor: tileBg, borderColor: tileBorder, color: textColor }}
                title={`${invName}: status ${statusVal}`}
              >
                <div style={{ fontSize: '0.65rem', color: 'var(--text-muted)' }}>
                  {invName.split('-')[0]}-{invName.split('-')[1]}
                </div>
                <strong>{invName.split('-')[2]}</strong>
                <div style={{ fontSize: '0.7rem', marginTop: '2px' }}>
                  {label}
                </div>
              </div>
            );
          })}
        </div>
      </div>

      {/* Daily Diagnostic Outages & Curtailments */}
      <div className="card">
        <h2 className="card-title">Dettaglio Diagnostico delle Perdite del Giorno {formattedTitleDate}</h2>
        {lossEvents.length === 0 ? (
          <p style={{ color: 'var(--success)', fontSize: '0.9rem' }}>
            Nessun evento critico registrato. Le perdite energetiche sono all'interno dei limiti nominali e non ci sono stati fermi inverter o limitazioni di rete.
          </p>
        ) : (
          <div style={{ maxHeight: '300px', overflowY: 'auto', paddingRight: '0.5rem' }}>
            <ul style={{ listStyleType: 'none', display: 'flex', flexDirection: 'column', gap: '0.65rem' }}>
              {lossEvents.map((evt, idx) => (
                <li key={idx} style={{ 
                  display: 'flex', 
                  alignItems: 'flex-start', 
                  gap: '0.5rem', 
                  padding: '0.65rem 1rem', 
                  backgroundColor: evt.type === 'curtailment' ? 'var(--warning-bg)' : 'var(--danger-bg)', 
                  borderLeft: `3px solid ${evt.type === 'curtailment' ? 'var(--warning)' : 'var(--danger)'}`,
                  borderRadius: '0.25rem',
                  fontSize: '0.875rem'
                }}>
                  <AlertTriangle 
                    size={16} 
                    color={evt.type === 'curtailment' ? 'var(--warning)' : 'var(--danger)'} 
                    style={{ marginTop: '0.1rem', flexShrink: 0 }} 
                  />
                  <div>
                    <span style={{ fontWeight: 600, color: 'var(--text-primary)' }}>
                      {evt.type === 'curtailment' ? 'LIMITAZIONE RETE' : 'FERMO INVERTER'}
                    </span>
                    <p style={{ color: 'var(--text-secondary)', marginTop: '0.15rem' }}>{evt.description}</p>
                  </div>
                </li>
              ))}
            </ul>
          </div>
        )}
      </div>
    </>
  );
}
