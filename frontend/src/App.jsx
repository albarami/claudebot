import React, { useState, useCallback, useRef, useEffect } from 'react'

const styles = {
  container: {
    minHeight: '100vh',
    background: 'linear-gradient(135deg, #1a1a2e 0%, #16213e 50%, #0f3460 100%)',
    color: '#fff',
    padding: '20px'
  },
  header: {
    textAlign: 'center',
    marginBottom: '30px',
    padding: '20px',
    background: 'rgba(255,255,255,0.05)',
    borderRadius: '16px',
    border: '1px solid rgba(255,255,255,0.1)'
  },
  title: {
    fontSize: '2.5rem',
    fontWeight: 'bold',
    marginBottom: '10px',
    background: 'linear-gradient(90deg, #00d4ff, #7b2cbf)',
    WebkitBackgroundClip: 'text',
    WebkitTextFillColor: 'transparent'
  },
  subtitle: {
    color: '#888',
    fontSize: '1rem'
  },
  main: {
    maxWidth: '1400px',
    margin: '0 auto',
    display: 'grid',
    gridTemplateColumns: '1fr 2fr',
    gap: '20px'
  },
  card: {
    background: 'rgba(255,255,255,0.05)',
    borderRadius: '16px',
    padding: '24px',
    border: '1px solid rgba(255,255,255,0.1)'
  },
  cardTitle: {
    fontSize: '1.25rem',
    fontWeight: '600',
    marginBottom: '16px',
    display: 'flex',
    alignItems: 'center',
    gap: '10px'
  },
  uploadArea: {
    border: '2px dashed rgba(255,255,255,0.2)',
    borderRadius: '12px',
    padding: '40px',
    textAlign: 'center',
    cursor: 'pointer',
    transition: 'all 0.3s',
    marginBottom: '20px'
  },
  button: {
    background: 'linear-gradient(90deg, #00d4ff, #7b2cbf)',
    color: '#fff',
    border: 'none',
    padding: '14px 28px',
    borderRadius: '8px',
    fontSize: '1rem',
    fontWeight: '600',
    cursor: 'pointer',
    width: '100%',
    marginTop: '16px',
    transition: 'transform 0.2s'
  },
  buttonDisabled: {
    background: '#444',
    cursor: 'not-allowed'
  },
  agentCard: {
    background: 'rgba(255,255,255,0.03)',
    borderRadius: '8px',
    padding: '12px',
    marginBottom: '10px',
    border: '1px solid rgba(255,255,255,0.05)'
  },
  agentName: {
    fontWeight: '600',
    marginBottom: '4px'
  },
  agentModel: {
    fontSize: '0.75rem',
    color: '#888'
  },
  logContainer: {
    height: '500px',
    overflowY: 'auto',
    background: '#0a0a0f',
    borderRadius: '8px',
    padding: '16px',
    fontFamily: 'monospace',
    fontSize: '0.85rem'
  },
  logEntry: {
    marginBottom: '8px',
    padding: '8px',
    borderRadius: '4px',
    background: 'rgba(255,255,255,0.02)'
  },
  progressBar: {
    height: '8px',
    background: 'rgba(255,255,255,0.1)',
    borderRadius: '4px',
    overflow: 'hidden',
    marginTop: '16px'
  },
  progressFill: {
    height: '100%',
    background: 'linear-gradient(90deg, #00d4ff, #7b2cbf)',
    transition: 'width 0.3s'
  },
  statusBadge: {
    display: 'inline-block',
    padding: '4px 12px',
    borderRadius: '20px',
    fontSize: '0.75rem',
    fontWeight: '600'
  },
  scoreCard: {
    textAlign: 'center',
    padding: '20px',
    background: 'rgba(0,212,255,0.1)',
    borderRadius: '12px',
    marginTop: '20px'
  },
  score: {
    fontSize: '3rem',
    fontWeight: 'bold',
    color: '#00d4ff'
  }
}

const agents = [
  { id: 'strategist', name: 'Strategist', model: 'Claude Sonnet 4.5', color: '#00d4ff' },
  { id: 'implementer', name: 'Implementer', model: 'Claude Opus 4.5', color: '#7b2cbf' },
  { id: 'qc_reviewer', name: 'QC Reviewer', model: 'Sonnet 4.5 + OpenAI 5.2', color: '#ff6b6b' },
  { id: 'auditor', name: 'Auditor', model: 'OpenAI 5.2', color: '#ffd93d' }
]

export default function App() {
  const [file, setFile] = useState(null)
  const [sessionId, setSessionId] = useState(null)
  const [status, setStatus] = useState('idle')
  const [progress, setProgress] = useState(0)
  const [logs, setLogs] = useState([])
  const [currentAgent, setCurrentAgent] = useState(null)
  const [result, setResult] = useState(null)
  const fileInputRef = useRef(null)
  const logContainerRef = useRef(null)

  useEffect(() => {
    if (logContainerRef.current) {
      logContainerRef.current.scrollTop = logContainerRef.current.scrollHeight
    }
  }, [logs])

  const addLog = (agent, message) => {
    const time = new Date().toLocaleTimeString()
    setLogs(prev => [...prev, { time, agent, message }])
  }

  const handleFileSelect = (e) => {
    const selected = e.target.files?.[0]
    if (selected && (selected.name.endsWith('.xlsx') || selected.name.endsWith('.xls'))) {
      setFile(selected)
      addLog('system', `Selected: ${selected.name}`)
    }
  }

  const handleUploadClick = () => {
    fileInputRef.current?.click()
  }

  const runAnalysis = async () => {
    if (!file) return

    setStatus('running')
    setProgress(0)
    setLogs([])
    setResult(null)

    try {
      addLog('system', 'Uploading file...')
      const formData = new FormData()
      formData.append('file', file)

      const uploadRes = await fetch('/api/upload', {
        method: 'POST',
        body: formData
      })

      if (!uploadRes.ok) throw new Error('Upload failed')
      const uploadData = await uploadRes.json()
      setSessionId(uploadData.session_id)
      addLog('system', `✓ Uploaded: ${uploadData.session_id}`)

      addLog('system', 'Starting LangGraph workflow...')
      const analyzeRes = await fetch('/api/analyze', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ session_id: uploadData.session_id })
      })

      if (!analyzeRes.ok) throw new Error('Failed to start analysis')
      addLog('system', '✓ Workflow started')

      let lastLogCount = 0
      while (true) {
        await new Promise(r => setTimeout(r, 2000))

        const statusRes = await fetch(`/api/status/${uploadData.session_id}`)
        if (!statusRes.ok) throw new Error('Status check failed')

        const data = await statusRes.json()
        setProgress(data.progress || 0)
        setCurrentAgent(data.current_agent)

        if (data.logs && data.logs.length > lastLogCount) {
          const newLogs = data.logs.slice(lastLogCount)
          for (const log of newLogs) {
            addLog(log.agent, log.message)
          }
          lastLogCount = data.logs.length
        }

        if (data.status === 'completed') {
          setResult({
            score: data.overall_score,
            certification: data.certification
          })
          addLog('system', `🏆 Complete! Score: ${data.overall_score?.toFixed(1)}%`)
          break
        }

        if (data.status === 'error' || data.status === 'failed') {
          throw new Error(data.errors?.join(', ') || 'Analysis failed')
        }
      }

      setStatus('completed')

    } catch (err) {
      setStatus('error')
      addLog('system', `❌ Error: ${err.message}`)
    }
  }

  const getAgentColor = (agentId) => {
    const agent = agents.find(a => a.id === agentId || agentId?.includes(a.id))
    return agent?.color || '#888'
  }

  const getStatusColor = () => {
    switch (status) {
      case 'running': return '#00d4ff'
      case 'completed': return '#4ade80'
      case 'error': return '#ef4444'
      default: return '#888'
    }
  }

  return (
    <div style={styles.container}>
      <header style={styles.header}>
        <h1 style={styles.title}>🎓 PhD Survey Analyzer</h1>
        <p style={styles.subtitle}>
          LangGraph Multi-Agent System • Sonnet 4.5 + Opus 4.5 + OpenAI 5.2
        </p>
      </header>

      <main style={styles.main}>
        <div>
          <div style={styles.card}>
            <h2 style={styles.cardTitle}>📁 Upload Survey</h2>
            
            <div 
              style={styles.uploadArea}
              onClick={handleUploadClick}
            >
              <input
                type="file"
                ref={fileInputRef}
                onChange={handleFileSelect}
                accept=".xlsx,.xls"
                style={{ display: 'none' }}
              />
              {file ? (
                <div>
                  <div style={{ fontSize: '2rem', marginBottom: '10px' }}>📊</div>
                  <div style={{ fontWeight: '600' }}>{file.name}</div>
                  <div style={{ color: '#888', fontSize: '0.85rem' }}>
                    {(file.size / 1024).toFixed(1)} KB
                  </div>
                </div>
              ) : (
                <div>
                  <div style={{ fontSize: '2rem', marginBottom: '10px' }}>📤</div>
                  <div>Click to upload Excel file</div>
                  <div style={{ color: '#888', fontSize: '0.85rem' }}>.xlsx or .xls</div>
                </div>
              )}
            </div>

            <button
              style={{
                ...styles.button,
                ...((!file || status === 'running') ? styles.buttonDisabled : {})
              }}
              onClick={runAnalysis}
              disabled={!file || status === 'running'}
            >
              {status === 'running' ? '⏳ Analyzing...' : '🚀 Run PhD-Level EDA'}
            </button>

            {status !== 'idle' && (
              <div style={styles.progressBar}>
                <div style={{ ...styles.progressFill, width: `${progress}%` }} />
              </div>
            )}

            {result && (
              <div style={styles.scoreCard}>
                <div style={{ color: '#888', marginBottom: '8px' }}>Quality Score</div>
                <div style={styles.score}>{result.score?.toFixed(1)}%</div>
                <div style={{
                  ...styles.statusBadge,
                  background: result.certification === 'PUBLICATION-READY' ? '#4ade80' : '#ffd93d',
                  color: '#000'
                }}>
                  {result.certification}
                </div>
              </div>
            )}
          </div>

          <div style={{ ...styles.card, marginTop: '20px' }}>
            <h2 style={styles.cardTitle}>🤖 Agents</h2>
            {agents.map(agent => (
              <div 
                key={agent.id} 
                style={{
                  ...styles.agentCard,
                  borderLeft: `3px solid ${agent.color}`,
                  opacity: currentAgent === agent.id ? 1 : 0.6
                }}
              >
                <div style={styles.agentName}>
                  {currentAgent === agent.id && '▶ '}{agent.name}
                </div>
                <div style={styles.agentModel}>{agent.model}</div>
              </div>
            ))}
          </div>
        </div>

        <div style={styles.card}>
          <h2 style={styles.cardTitle}>
            📋 Execution Log
            <span style={{
              ...styles.statusBadge,
              background: getStatusColor(),
              marginLeft: 'auto'
            }}>
              {status.toUpperCase()}
            </span>
          </h2>
          
          <div style={styles.logContainer} ref={logContainerRef}>
            {logs.length === 0 ? (
              <div style={{ color: '#666', textAlign: 'center', marginTop: '200px' }}>
                Upload a survey and click "Run PhD-Level EDA" to start
              </div>
            ) : (
              logs.map((log, i) => (
                <div key={i} style={styles.logEntry}>
                  <span style={{ color: '#666' }}>[{log.time}]</span>
                  <span style={{ 
                    color: getAgentColor(log.agent),
                    fontWeight: '600',
                    marginLeft: '8px'
                  }}>
                    {log.agent}:
                  </span>
                  <span style={{ marginLeft: '8px' }}>{log.message}</span>
                </div>
              ))
            )}
          </div>
        </div>
      </main>
    </div>
  )
}
