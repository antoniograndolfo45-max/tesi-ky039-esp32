#include <Arduino.h>
#define PIN_PPG  34
#define ALPHA    0.75f
#define TS_MS    20          
void setup(){
  Serial.begin(115200);
  delay(100);
  analogReadResolution(12);                 
  analogSetPinAttenuation(PIN_PPG, ADC_11db); 
}
float readsensor(){
  static float last=0.0f;
  float sum=0.0f;

  for(int i= 0; i<20; i++){
    sum += analogRead(PIN_PPG);
    delay(1);
  };
  last = ALPHA * (sum / 20.0f) + (1.0f - ALPHA) * last;
  return last;           
}
// ================== RILEVAMENTO BATTITI ==================
// Rimuoviamo la baseline, stimiamo il rumore e applichiamo soglia con isteresi
static const float BASELINE_ALPHA = 0.01f;    
static const float NOISE_ALPHA    = 0.10f;   
static const float K_NOISE        = 1.5f;   
static const float MIN_AMP        = 4.0f;     
static const float HYST_FRAC      = 0.45f;    
static const unsigned long REF_BASE_MS = 240; 
static const unsigned long IBI_MIN = 300;     // 30 BPM
static const unsigned long IBI_MAX = 2000;    // 200 BPM
float baseline = 0.0f, noiseEMA = 0.0f, prevY = 0.0f;
bool  armed = true;
unsigned long lastBeatMs = 0, lastIBI = 800, lastToggleMs = 0;
float bpmToSend = 0.0f, lastGoodBPM = 0.0f;
int   contatore = 0; 

// ================== ORTOSTATICO ==================
enum State { IDLE, BASELINE, POST_STAND, DONE };
State state = IDLE;
unsigned long t_startBaseline=0;
unsigned long t_stand=0;
const int MAX_SEC = 240;
float hrSeries[MAX_SEC];
unsigned long tsSeries[MAX_SEC];
int hrIdx=0, hrCount=0;
bool baselineDone=false, peakDone=false, recovDone=false;
float HR_baseline=0, HR_peak=0, dHR=0, HR_recov60=0;
unsigned long t_peak_ms=0;

// ================== BUFFER HR 1 Hz ==================
void storeHR1Hz(){
  if (bpmToSend > 0 && bpmToSend < 220) {
    tsSeries[hrIdx] = millis();
    hrSeries[hrIdx] = bpmToSend;
    hrIdx = (hrIdx + 1) % MAX_SEC;
    hrCount = min(hrCount+1, MAX_SEC);
  }
}
bool avgHRwindow(unsigned long t0, unsigned long t1, float &avg){
  double s=0; int n=0;
  for (int i=0;i<hrCount;i++){
    int idx = (hrIdx - 1 - i + MAX_SEC) % MAX_SEC;
    unsigned long t = tsSeries[idx];
    if (t >= t0 && t < t1){ s += hrSeries[idx]; n++; }
  }
  if (n >= 5){ avg = (float)(s / n); return true; }
  return false;
}
bool maxHRwindow(unsigned long t0, unsigned long t1, float &mx, unsigned long &t_at){
  bool ok=false; mx=-1; t_at=0;
  for (int i=0;i<hrCount;i++){
    int idx = (hrIdx - 1 - i + MAX_SEC) % MAX_SEC;
    unsigned long t = tsSeries[idx];
    float v = hrSeries[idx];
    if (t >= t0 && t < t1){
      if (!ok || v > mx){ mx = v; t_at = t; ok=true; }
    }
  }
  return ok;
}
// ================== TICK 1 Hz di metriche ortostatico ==================
unsigned long nextTick = 0;
void tick1Hz(){
  storeHR1Hz();

  unsigned long now = millis();
  if (state == POST_STAND){
    // media 30 s prima di STAND
    if (!baselineDone && (now > t_stand)) {
      float avg;
      if (avgHRwindow(t_stand - 30000UL, t_stand, avg)) {
        HR_baseline = avg; baselineDone = true;
      }
    }
    // picco max 0..30 s dopo STAND
    if (!peakDone && (now - t_stand >= 30000UL)) {
      float mx; unsigned long t_at;
      if (maxHRwindow(t_stand, t_stand + 30000UL, mx, t_at)) {
        HR_peak = mx; t_peak_ms = (t_at > t_stand) ? (t_at - t_stand) : 0; peakDone = true;
      }
    }
    // media 60..120 s dopo STAND
    if (!recovDone && (now - t_stand >= 120000UL)) {
      float avg;
      if (avgHRwindow(t_stand + 60000UL, t_stand + 120000UL, avg)) {
        HR_recov60 = avg; recovDone = true;
      }
    }
    // stampa finale
    if (baselineDone && peakDone && recovDone) {
      dHR = HR_peak - HR_baseline;
      Serial.print("METRICS ");
      Serial.print("baseline="); Serial.print(HR_baseline, 1);
      Serial.print(" peak=");    Serial.print(HR_peak, 1);
      Serial.print(" dHR=");     Serial.print(dHR, 1);
      Serial.print(" tpeak=");   Serial.print((float)t_peak_ms/1000.0f, 1);
      Serial.print(" recov60="); Serial.println(HR_recov60, 1);
      state = DONE;
    }
  }
}

// ================== COMANDI SERIALI ==================
String cmdBuf;
void handleCommand(const String& cmd){
  if (cmd == "CMD:RESET"){
    state = IDLE;
    baselineDone = peakDone = recovDone = false;
    HR_baseline = HR_peak = dHR = HR_recov60 = 0; t_peak_ms = 0;
    hrIdx=0; hrCount=0;
    contatore = 0; lastBeatMs = 0; lastToggleMs = 0; bpmToSend = 0; lastGoodBPM = 0;
    baseline = 0; noiseEMA = 0; prevY = 0; armed = true; lastIBI = 800;
    Serial.println("ACK:RESET");
  } else if (cmd == "CMD:START"){
    state = BASELINE;
    baselineDone = peakDone = recovDone = false;
    HR_baseline = HR_peak = dHR = HR_recov60 = 0; t_peak_ms = 0;
    hrIdx=0; hrCount=0;
    contatore = 0; lastBeatMs = 0; lastToggleMs = 0; bpmToSend = 0; lastGoodBPM = 0;
    baseline = 0; noiseEMA = 0; prevY = 0; armed = true; lastIBI = 800;
    t_startBaseline = millis();
    Serial.println("ACK:START");
  } else if (cmd == "CMD:STAND"){
    if (state == BASELINE){
      state = POST_STAND;
      t_stand = millis();
      Serial.println("ACK:STAND");
    }
  }
}
void pollSerial(){
  while (Serial.available()){
    char c = (char)Serial.read();
    if (c == '\n' || c == '\r'){
      if (cmdBuf.length()){
        handleCommand(cmdBuf);
        cmdBuf = "";
      }
    } else {
      cmdBuf += c;
    }
  }
}
void loop(){
  unsigned long now = millis();

  // campionamento a ~50 Hz sincronizzato
  static unsigned long nextSample = 0;
  if ((long)(now - nextSample) < 0){
    pollSerial();
    return;
  }
  nextSample = now + TS_MS;

  // 1) acquisizione  
  float y = readsensor();

  // 2) rimozione DC + stima rumore
  float dy = y - prevY;
  if (baseline == 0.0f) baseline = y;
  baseline += BASELINE_ALPHA * (y - baseline);
  noiseEMA += NOISE_ALPHA    * (fabsf(dy) - noiseEMA);
  prevY = y;

  float y_ac = y - baseline;          

  // 3) soglie con isteresi
  float sigma = (noiseEMA > 0.5f) ? noiseEMA : 0.5f;
  float hi = max(MIN_AMP, K_NOISE * sigma);
  float lo = max(0.5f, hi * (1.0f - HYST_FRAC));

  // 4) refractory adattivo
  unsigned long refMs = max(REF_BASE_MS, (unsigned long)(0.45f * (float)lastIBI));

  // 5) trigger (solo fronte di salita)
  if (armed && y_ac > hi && (now - lastToggleMs) >= refMs){
    lastToggleMs = now;

    if (lastBeatMs > 0){
      unsigned long ibi = now - lastBeatMs;
      if (ibi >= IBI_MIN && ibi <= IBI_MAX){
        lastIBI = ibi;
        float bpm = 60000.0f / (float)ibi;
        bpmToSend = 0.6f * bpmToSend + 0.4f * bpm;
        lastGoodBPM = bpmToSend;

        Serial.print("BPM: ");
        Serial.println(bpmToSend, 1);
      }
    }
    lastBeatMs = now;
    armed = false;
  } else if (!armed && y_ac < lo){
    armed = true;
  }
  // 6) metriche ortostatico 1 Hz
  static unsigned long nextTick = millis() + 1000; 
  if ((long)(now - nextTick) >= 0){
    tick1Hz();
    nextTick += 1000;
  }

  pollSerial();
}