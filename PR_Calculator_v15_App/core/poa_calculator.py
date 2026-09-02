import numpy as np

def compute_poa_series(poa1_series, poa3_series, threshold=50.0, method='condmax'):
    """
    Computes effective POA series using Conditional MAX, Average, TX1, or TX3 method.
    Conditional MAX: if both sensors > threshold and agree within 10%, max is used;
    otherwise falling back to valid sensor.
    """
    n = len(poa1_series)
    res_w = np.zeros(n)
    res_kwh = np.zeros(n)
    
    for i in range(n):
        p1 = float(poa1_series[i]) if i < len(poa1_series) else 0.0
        p3 = float(poa3_series[i]) if i < len(poa3_series) else 0.0
        
        if p1 <= 0 and p3 > 0:
            p1 = p3
        elif p3 <= 0 and p1 > 0:
            p3 = p1
            
        if method == 'tx1':
            val_w = p1
        elif method == 'tx3':
            val_w = p3
        elif method == 'avg':
            val_w = (p1 + p3) / 2.0
        else: # condmax
            if p1 > threshold and p3 > threshold:
                avg = (p1 + p3) / 2.0
                if abs(p1 - p3) / avg <= 0.10:
                    val_w = max(p1, p3)
                else:
                    val_w = p1
            elif p1 > threshold:
                val_w = p1
            elif p3 > threshold:
                val_w = p3
            else:
                val_w = max(p1, p3)
                
        res_w[i] = val_w if val_w > threshold else 0.0
        res_kwh[i] = (val_w / 4000.0) if val_w > threshold else 0.0
        
    return res_w, res_kwh
