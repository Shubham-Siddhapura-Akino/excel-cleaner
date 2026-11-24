'use client';

import React, { useState, useRef, useEffect } from 'react';
import {
  FileSpreadsheet,
  Upload,
  Download,
  Trash2,
  CheckCircle2,
  AlertCircle,
  Clock,
  Sparkles,
  Layers,
  Archive,
  RefreshCw,
  Eye,
  Sliders,
  History,
  X,
  FileCheck
} from 'lucide-react';

export default function Home() {
  // State
  const [attachedFiles, setAttachedFiles] = useState([]);
  const [isDragging, setIsDragging] = useState(false);
  const [isProcessing, setIsProcessing] = useState(false);
  const [processingStatus, setProcessingStatus] = useState({ current: 0, total: 0, filename: '', percent: 0 });
  const [processedResults, setProcessedResults] = useState([]);
  const [showResults, setShowResults] = useState(false);
  const [openDiffIndex, setOpenDiffIndex] = useState(null);

  // Cleaning options
  const [optQuotes, setOptQuotes] = useState(true);
  const [optNormalize, setOptNormalize] = useState(true);
  const [optTrim, setOptTrim] = useState(true);
  const [optPruneRows, setOptPruneRows] = useState(true);
  const [optAutoDownload, setOptAutoDownload] = useState(true);

  // History modal
  const [showHistory, setShowHistory] = useState(false);
  const [historyFiles, setHistoryFiles] = useState([]);
  const [loadingHistory, setLoadingHistory] = useState(false);

  // Toast
  const [toast, setToast] = useState(null);

  const fileInputRef = useRef(null);

  const showToastMsg = (message, isError = false) => {
    setToast({ message, isError });
    setTimeout(() => {
      setToast(null);
    }, 4500);
  };

  const formatBytes = (bytes) => {
    if (!bytes || bytes === 0) return '0 B';
    const k = 1024;
    const sizes = ['B', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
  };

  // Add files to queue
  const handleAddFiles = (files) => {
    if (!files || files.length === 0) return;
    const validExts = ['.xlsx', '.xlsm', '.xls'];
    const newFiles = [];

    Array.from(files).forEach((file) => {
      const ext = file.name.substring(file.name.lastIndexOf('.')).toLowerCase();
      if (validExts.includes(ext)) {
        const exists = attachedFiles.some((f) => f.name === file.name && f.size === file.size);
        if (!exists) {
          newFiles.push(file);
        }
      }
    });

    if (newFiles.length === 0 && files.length > 0) {
      showToastMsg('Only valid Excel files (.xlsx, .xlsm, .xls) are accepted (or file already added)', true);
    } else {
      setAttachedFiles((prev) => [...prev, ...newFiles]);
    }
  };

  const removeFile = (index) => {
    setAttachedFiles((prev) => prev.filter((_, i) => i !== index));
  };

  const clearAllFiles = () => {
    setAttachedFiles([]);
    if (fileInputRef.current) fileInputRef.current.value = '';
  };

  // Drag and Drop Handlers
  const onDragOver = (e) => {
    e.preventDefault();
    setIsDragging(true);
  };

  const onDragLeave = (e) => {
    e.preventDefault();
    setIsDragging(false);
  };

  const onDrop = (e) => {
    e.preventDefault();
    setIsDragging(false);
    if (e.dataTransfer.files) {
      handleAddFiles(e.dataTransfer.files);
    }
  };

  // Batch Cleaning Execution
  const handleCleanAll = async () => {
    if (attachedFiles.length === 0) return;

    setIsProcessing(true);
    setShowResults(false);
    setProcessedResults([]);

    const totalFiles = attachedFiles.length;
    const results = [];
    const startTime = Date.now();

    const options = {
      removeQuotes: optQuotes,
      normalizeWhitespace: optNormalize,
      trimSpaces: optTrim,
      pruneEmptyRows: optPruneRows
    };

    for (let i = 0; i < totalFiles; i++) {
      const file = attachedFiles[i];
      const percent = Math.round(((i + 1) / totalFiles) * 100);

      setProcessingStatus({
        current: i + 1,
        total: totalFiles,
        filename: file.name,
        percent
      });

      try {
        const res = await fetch('/api/clean', {
          method: 'POST',
          headers: {
            'Content-Type': 'application/octet-stream',
            'X-Filename': encodeURIComponent(file.name),
            'X-Options': encodeURIComponent(JSON.stringify(options))
          },
          body: file
        });

        const data = await res.json();
        if (!res.ok || !data.success) {
          throw new Error(data.error || 'Cleaning failed');
        }

        results.push({
          file,
          success: true,
          data
        });
      } catch (err) {
        results.push({
          file,
          success: false,
          error: err.message
        });
      }
    }

    const totalDuration = Date.now() - startTime;
    setIsProcessing(false);
    setProcessedResults(results);
    setShowResults(true);

    const successful = results.filter((r) => r.success);
    const allDownloadIds = successful.map((r) => r.data.downloadId);

    // Auto download if enabled
    if (optAutoDownload) {
      if (successful.length === 1) {
        triggerDownload(successful[0].data.downloadUrl, successful[0].data.cleanedFilename);
      } else if (successful.length > 1) {
        triggerDownload(`/api/download-zip?ids=${allDownloadIds.join(',')}`, 'cleaned_workbooks.zip');
      }
    }

    const aggModified = successful.reduce((sum, r) => sum + r.data.stats.modifiedCells, 0);
    showToastMsg(`Cleaned ${aggModified.toLocaleString()} cells across ${successful.length} file(s)!`);
  };

  const triggerDownload = (url, filename) => {
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
  };

  // Load history from API
  const fetchHistory = async () => {
    setLoadingHistory(true);
    setShowHistory(true);
    try {
      const res = await fetch('/api/history');
      const data = await res.json();
      setHistoryFiles(data.files || []);
    } catch (err) {
      setHistoryFiles([]);
    } finally {
      setLoadingHistory(false);
    }
  };

  // Aggregate stats calculations
  const successfulResults = processedResults.filter((r) => r.success);
  const aggSheets = successfulResults.reduce((sum, r) => sum + r.data.stats.sheetsProcessed, 0);
  const aggRows = successfulResults.reduce((sum, r) => sum + r.data.stats.rowsProcessed, 0);
  const aggCells = successfulResults.reduce((sum, r) => sum + r.data.stats.cellsVisited, 0);
  const aggModified = successfulResults.reduce((sum, r) => sum + r.data.stats.modifiedCells, 0);
  const allDownloadIds = successfulResults.map((r) => r.data.downloadId);

  return (
    <div className="relative min-h-screen flex flex-col justify-between">
      {/* Background Decorative Gradients */}
      <div className="fixed inset-0 pointer-events-none overflow-hidden z-0">
        <div className="absolute -top-40 left-1/2 -translate-x-1/2 w-[1000px] h-[500px] bg-gradient-to-b from-emerald-500/10 via-teal-500/5 to-transparent blur-3xl"></div>
        <div className="absolute top-1/3 -left-48 w-96 h-96 bg-blue-500/5 rounded-full blur-3xl"></div>
        <div className="absolute top-2/3 -right-48 w-96 h-96 bg-emerald-500/5 rounded-full blur-3xl"></div>
      </div>

      {/* Header */}
      <header className="relative z-10 border-b border-slate-800/80 bg-slate-900/60 backdrop-blur-md sticky top-0">
        <div className="max-w-6xl mx-auto px-4 sm:px-6 h-16 flex items-center justify-between">
          <div className="flex items-center space-x-3">
            <div className="w-10 h-10 rounded-xl bg-gradient-to-tr from-emerald-600 to-teal-400 flex items-center justify-center shadow-lg shadow-emerald-500/20">
              <FileSpreadsheet className="w-6 h-6 text-white" />
            </div>
            <div>
              <div className="flex items-center space-x-2">
                <span className="font-bold text-lg tracking-tight text-white">Excel Cleaner</span>
                <span className="text-xs font-semibold px-2 py-0.5 rounded-full bg-emerald-500/10 text-emerald-400 border border-emerald-500/20">
                  Next.js + Fastify
                </span>
              </div>
              <p className="text-xs text-slate-400">Vercel Ready • Multi-Workbook Batch Cleaner</p>
            </div>
          </div>

          <div className="flex items-center space-x-4">
            <div className="hidden sm:flex items-center space-x-2 text-xs font-medium text-emerald-400 bg-emerald-500/10 px-3 py-1.5 rounded-full border border-emerald-500/20">
              <span className="w-2 h-2 rounded-full bg-emerald-400 animate-ping"></span>
              <span>Fastify Engine Active</span>
            </div>
            <button
              onClick={fetchHistory}
              className="text-xs font-medium text-slate-300 hover:text-white bg-slate-800 hover:bg-slate-700 transition px-3 py-1.5 rounded-lg border border-slate-700 flex items-center space-x-1.5"
            >
              <History className="w-3.5 h-3.5 text-slate-400" />
              <span>History</span>
            </button>
          </div>
        </div>
      </header>

      {/* Main Container */}
      <main className="relative z-10 flex-1 max-w-5xl w-full mx-auto px-4 sm:px-6 py-8 sm:py-12 flex flex-col space-y-8">
        {/* Hero Title */}
        <div className="text-center space-y-3 max-w-2xl mx-auto">
          <h1 className="text-3xl sm:text-4xl font-extrabold tracking-tight text-white">
            Clean & Normalize Excel Workbooks
          </h1>
          <p className="text-sm sm:text-base text-slate-400 leading-relaxed">
            Attach multiple spreadsheets at once. Strip double quotes, normalize whitespace, remove invisible Unicode characters, and download all files at once as a ZIP bundle or individually.
          </p>
        </div>

        {/* Feature Highlights */}
        <div className="grid grid-cols-1 sm:grid-cols-3 gap-3 text-xs text-slate-400">
          <div className="flex items-center space-x-2 p-3 rounded-xl bg-slate-900/50 border border-slate-800/80">
            <Layers className="w-4 h-4 text-emerald-400 shrink-0" />
            <span>Multi-File Attachment & Queue</span>
          </div>
          <div className="flex items-center space-x-2 p-3 rounded-xl bg-slate-900/50 border border-slate-800/80">
            <Archive className="w-4 h-4 text-emerald-400 shrink-0" />
            <span>Download All At Once (.ZIP or Single)</span>
          </div>
          <div className="flex items-center space-x-2 p-3 rounded-xl bg-slate-900/50 border border-slate-800/80">
            <FileCheck className="w-4 h-4 text-emerald-400 shrink-0" />
            <span>100% Formulas, Styles & Formats Preserved</span>
          </div>
        </div>

        {/* Upload Card */}
        {!isProcessing && !showResults && (
          <div className="glass-card rounded-2xl p-6 sm:p-8 shadow-2xl relative overflow-hidden transition-all duration-300">
            {/* Drag & Drop Area */}
            <div
              onClick={() => fileInputRef.current?.click()}
              onDragOver={onDragOver}
              onDragLeave={onDragLeave}
              onDrop={onDrop}
              className={`border-2 border-dashed rounded-xl p-8 sm:p-12 text-center cursor-pointer transition-all duration-200 group flex flex-col items-center justify-center space-y-4 ${
                isDragging
                  ? 'border-emerald-500 bg-slate-800/60'
                  : 'border-slate-700 hover:border-emerald-500 bg-slate-900/40 hover:bg-slate-800/40'
              }`}
            >
              <input
                ref={fileInputRef}
                type="file"
                className="hidden"
                multiple
                accept=".xlsx,.xlsm,.xls"
                onChange={(e) => {
                  handleAddFiles(e.target.files);
                  e.target.value = '';
                }}
              />

              <div className="w-16 h-16 rounded-2xl bg-slate-800 group-hover:bg-emerald-500/20 text-slate-400 group-hover:text-emerald-400 flex items-center justify-center transition-all duration-200 group-hover:scale-110 shadow-inner">
                <Upload className="w-8 h-8" />
              </div>

              <div className="space-y-1">
                <p className="text-base font-semibold text-slate-200 group-hover:text-white transition">
                  Attach one or multiple Excel files here, or <span className="text-emerald-400 underline decoration-emerald-500/50 underline-offset-2">browse files</span>
                </p>
                <p className="text-xs text-slate-400">Hold Ctrl / Cmd to select multiple workbooks or drag them together</p>
              </div>

              <div className="flex items-center space-x-2 text-xs text-slate-500 pt-2">
                <span className="px-2 py-0.5 rounded bg-slate-800 text-slate-400 font-mono">.xlsx</span>
                <span className="px-2 py-0.5 rounded bg-slate-800 text-slate-400 font-mono">.xlsm</span>
                <span>• Fastify Batch Engine</span>
              </div>
            </div>

            {/* Attached Files Queue */}
            {attachedFiles.length > 0 && (
              <div className="mt-6 space-y-3">
                <div className="flex items-center justify-between pb-2 border-b border-slate-800">
                  <div className="flex items-center space-x-2">
                    <span className="text-xs font-bold uppercase tracking-wider text-slate-300">Attached Workbooks</span>
                    <span className="text-xs px-2 py-0.5 rounded-full bg-emerald-500/20 text-emerald-400 font-bold">
                      {attachedFiles.length}
                    </span>
                  </div>
                  <div className="flex items-center space-x-2">
                    <button
                      type="button"
                      onClick={() => fileInputRef.current?.click()}
                      className="text-xs text-emerald-400 hover:text-emerald-300 bg-emerald-500/10 hover:bg-emerald-500/20 border border-emerald-500/30 px-3 py-1 rounded-lg transition flex items-center space-x-1"
                    >
                      <span>Attach More</span>
                    </button>
                    <button
                      type="button"
                      onClick={clearAllFiles}
                      className="text-xs text-slate-400 hover:text-rose-400 hover:bg-rose-500/10 px-2 py-1 rounded-lg transition"
                    >
                      Clear All
                    </button>
                  </div>
                </div>

                <div className="max-h-60 overflow-y-auto space-y-2 pr-1">
                  {attachedFiles.map((file, idx) => (
                    <div
                      key={`${file.name}-${idx}`}
                      className="flex items-center justify-between p-3 rounded-xl bg-slate-900/60 border border-slate-800 text-xs hover:border-slate-700 transition"
                    >
                      <div className="flex items-center space-x-3 truncate">
                        <div className="w-8 h-8 rounded-lg bg-emerald-500/10 text-emerald-400 flex items-center justify-center shrink-0">
                          <FileSpreadsheet className="w-4 h-4" />
                        </div>
                        <div className="truncate">
                          <div className="font-medium text-slate-200 truncate font-mono">{file.name}</div>
                          <div className="text-[11px] text-slate-400 font-mono">{formatBytes(file.size)}</div>
                        </div>
                      </div>
                      <button
                        type="button"
                        onClick={() => removeFile(idx)}
                        className="p-1.5 text-slate-400 hover:text-rose-400 hover:bg-rose-500/10 rounded-lg transition"
                        title="Remove file"
                      >
                        <Trash2 className="w-4 h-4" />
                      </button>
                    </div>
                  ))}
                </div>
              </div>
            )}

            {/* Configurable Cleaning Options */}
            <div className="mt-6 border-t border-slate-800 pt-5">
              <div className="flex items-center justify-between mb-4">
                <div className="flex items-center space-x-2 text-xs font-semibold uppercase tracking-wider text-slate-400">
                  <Sliders className="w-4 h-4 text-emerald-400" />
                  <span>Cleaning Options</span>
                </div>
                <span className="text-xs text-emerald-400 font-medium">Standard Strict Normalization</span>
              </div>

              <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 text-xs">
                <label className="flex items-start space-x-3 p-3 rounded-lg bg-slate-900/60 border border-slate-800 cursor-pointer hover:border-slate-700 transition">
                  <input
                    type="checkbox"
                    checked={optQuotes}
                    onChange={(e) => setOptQuotes(e.target.checked)}
                    className="mt-0.5 rounded text-emerald-500 focus:ring-emerald-500 bg-slate-800 border-slate-700"
                  />
                  <div>
                    <span className="font-medium text-slate-200 block">Remove Double Quotes</span>
                    <span className="text-slate-400 text-[11px]">Deletes all literal quotes (&quot;) from text values</span>
                  </div>
                </label>

                <label className="flex items-start space-x-3 p-3 rounded-lg bg-slate-900/60 border border-slate-800 cursor-pointer hover:border-slate-700 transition">
                  <input
                    type="checkbox"
                    checked={optNormalize}
                    onChange={(e) => setOptNormalize(e.target.checked)}
                    className="mt-0.5 rounded text-emerald-500 focus:ring-emerald-500 bg-slate-800 border-slate-700"
                  />
                  <div>
                    <span className="font-medium text-slate-200 block">Normalize Whitespace</span>
                    <span className="text-slate-400 text-[11px]">Replaces tabs, newlines, multi-spaces & invisible chars with 1 space</span>
                  </div>
                </label>

                <label className="flex items-start space-x-3 p-3 rounded-lg bg-slate-900/60 border border-slate-800 cursor-pointer hover:border-slate-700 transition">
                  <input
                    type="checkbox"
                    checked={optTrim}
                    onChange={(e) => setOptTrim(e.target.checked)}
                    className="mt-0.5 rounded text-emerald-500 focus:ring-emerald-500 bg-slate-800 border-slate-700"
                  />
                  <div>
                    <span className="font-medium text-slate-200 block">Trim Leading & Trailing</span>
                    <span className="text-slate-400 text-[11px]">Strips outer blank spaces from all cell values</span>
                  </div>
                </label>

                <label className="flex items-start space-x-3 p-3 rounded-lg bg-slate-900/60 border border-slate-800 cursor-pointer hover:border-slate-700 transition">
                  <input
                    type="checkbox"
                    checked={optPruneRows}
                    onChange={(e) => setOptPruneRows(e.target.checked)}
                    className="mt-0.5 rounded text-emerald-500 focus:ring-emerald-500 bg-slate-800 border-slate-700"
                  />
                  <div>
                    <span className="font-medium text-slate-200 block">Prune Empty Trailing Rows</span>
                    <span className="text-slate-400 text-[11px]">Prevents bloated empty row XML tags</span>
                  </div>
                </label>
              </div>

              <div className="mt-3 flex items-center justify-between text-xs text-slate-400">
                <label className="flex items-center space-x-2 cursor-pointer">
                  <input
                    type="checkbox"
                    checked={optAutoDownload}
                    onChange={(e) => setOptAutoDownload(e.target.checked)}
                    className="rounded text-emerald-500 focus:ring-emerald-500 bg-slate-800 border-slate-700"
                  />
                  <span>Auto-download all cleaned files (as ZIP bundle or direct) when finished</span>
                </label>
              </div>
            </div>

            {/* Action CTA Button */}
            <div className="mt-6 flex flex-col sm:flex-row items-center gap-3">
              <button
                type="button"
                disabled={attachedFiles.length === 0}
                onClick={handleCleanAll}
                className="w-full sm:flex-1 py-3.5 px-6 rounded-xl font-semibold text-white bg-emerald-600 hover:bg-emerald-500 disabled:opacity-50 disabled:cursor-not-allowed shadow-lg shadow-emerald-600/30 transition-all duration-200 flex items-center justify-center space-x-2 text-sm"
              >
                <Sparkles className="w-4 h-4" />
                <span>
                  {attachedFiles.length === 0
                    ? 'Attach Excel Files to Clean'
                    : attachedFiles.length === 1
                    ? 'Clean 1 Attached Workbook'
                    : `Clean All ${attachedFiles.length} Attached Workbooks`}
                </span>
              </button>
            </div>
          </div>
        )}

        {/* Processing State */}
        {isProcessing && (
          <div className="glass-card rounded-2xl p-8 text-center space-y-5 animate-pulse-subtle">
            <div className="w-14 h-14 mx-auto rounded-2xl bg-emerald-500/10 border border-emerald-500/20 text-emerald-400 flex items-center justify-center">
              <RefreshCw className="w-7 h-7 animate-spin" />
            </div>
            <div className="space-y-1">
              <h3 className="text-lg font-bold text-white">
                Cleaning Workbook {processingStatus.current} of {processingStatus.total}...
              </h3>
              <p className="text-xs text-slate-400 font-mono">
                Normalizing &quot;{processingStatus.filename}&quot;...
              </p>
            </div>
            <div className="w-full bg-slate-800 rounded-full h-2 max-w-md mx-auto overflow-hidden">
              <div
                className="bg-gradient-to-r from-emerald-500 to-teal-400 h-2 rounded-full transition-all duration-300"
                style={{ width: `${processingStatus.percent}%` }}
              ></div>
            </div>
          </div>
        )}

        {/* Results Screen */}
        {showResults && (
          <div className="space-y-6">
            {/* Success Callout Banner */}
            <div className="glass-card rounded-2xl p-6 sm:p-8 border border-emerald-500/30 bg-gradient-to-r from-emerald-950/40 via-slate-900/60 to-slate-900/60 shadow-2xl flex flex-col sm:flex-row items-center justify-between gap-6">
              <div className="flex items-center space-x-4">
                <div className="w-12 h-12 rounded-2xl bg-emerald-500 text-white flex items-center justify-center shadow-lg shadow-emerald-500/30 shrink-0">
                  <CheckCircle2 className="w-6 h-6" />
                </div>
                <div>
                  <div className="flex items-center space-x-2">
                    <h2 className="text-lg font-bold text-white">
                      {successfulResults.length === 1
                        ? 'Workbook Cleaned Successfully!'
                        : `All ${successfulResults.length} Workbooks Cleaned Successfully!`}
                    </h2>
                  </div>
                  <p className="text-xs text-slate-400 font-mono mt-0.5">
                    {successfulResults.length} of {processedResults.length} files processed without errors
                  </p>
                </div>
              </div>

              <div className="flex flex-col sm:flex-row items-center gap-3 w-full sm:w-auto">
                {successfulResults.length === 1 ? (
                  <a
                    href={successfulResults[0].data.downloadUrl}
                    download={successfulResults[0].data.cleanedFilename}
                    className="w-full sm:w-auto py-3.5 px-6 rounded-xl font-bold text-white bg-gradient-to-r from-emerald-600 to-teal-500 hover:from-emerald-500 hover:to-teal-400 shadow-xl shadow-emerald-600/30 transition-all flex items-center justify-center space-x-2 text-sm cursor-pointer"
                  >
                    <Download className="w-5 h-5" />
                    <span>Download Cleaned File</span>
                  </a>
                ) : (
                  <a
                    href={`/api/download-zip?ids=${allDownloadIds.join(',')}`}
                    download="cleaned_workbooks.zip"
                    className="w-full sm:w-auto py-3.5 px-6 rounded-xl font-bold text-white bg-gradient-to-r from-emerald-600 via-emerald-500 to-teal-500 hover:from-emerald-500 hover:to-teal-400 shadow-xl shadow-emerald-600/30 transition-all flex items-center justify-center space-x-2 text-sm cursor-pointer"
                  >
                    <Download className="w-5 h-5" />
                    <span>Download All as ZIP ({successfulResults.length} Files)</span>
                  </a>
                )}

                <button
                  onClick={() => {
                    setShowResults(false);
                    setAttachedFiles([]);
                    setProcessedResults([]);
                  }}
                  className="w-full sm:w-auto py-3 px-4 rounded-xl text-slate-300 hover:text-white bg-slate-800 hover:bg-slate-700 border border-slate-700 text-sm font-medium transition"
                >
                  Clean More Files
                </button>
              </div>
            </div>

            {/* Aggregate Dashboard Metrics */}
            <div className="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-5 gap-3">
              <div className="glass-card rounded-xl p-4 text-center">
                <div className="text-[11px] font-medium text-slate-400 uppercase tracking-wider">Workbooks</div>
                <div className="text-xl font-extrabold text-white mt-1 font-mono">{successfulResults.length}</div>
              </div>
              <div className="glass-card rounded-xl p-4 text-center">
                <div className="text-[11px] font-medium text-slate-400 uppercase tracking-wider">Total Sheets</div>
                <div className="text-xl font-extrabold text-white mt-1 font-mono">{aggSheets.toLocaleString()}</div>
              </div>
              <div className="glass-card rounded-xl p-4 text-center">
                <div className="text-[11px] font-medium text-slate-400 uppercase tracking-wider">Total Rows</div>
                <div className="text-xl font-extrabold text-white mt-1 font-mono">{aggRows.toLocaleString()}</div>
              </div>
              <div className="glass-card rounded-xl p-4 text-center">
                <div className="text-[11px] font-medium text-slate-400 uppercase tracking-wider">Cells Visited</div>
                <div className="text-xl font-extrabold text-white mt-1 font-mono">{aggCells.toLocaleString()}</div>
              </div>
              <div className="glass-card rounded-xl p-4 text-center border-emerald-500/30 bg-emerald-950/20">
                <div className="text-[11px] font-medium text-emerald-400 uppercase tracking-wider">Modified Cells</div>
                <div className="text-xl font-extrabold text-emerald-400 mt-1 font-mono">{aggModified.toLocaleString()}</div>
              </div>
            </div>

            {/* Per-File Breakdown Cards */}
            <div className="space-y-4">
              <div className="flex items-center justify-between">
                <h3 className="text-sm font-bold uppercase tracking-wider text-slate-300">Cleaned Files Breakdown</h3>
                <span className="text-xs text-slate-400">Download files individually or use the master button above</span>
              </div>

              <div className="space-y-3">
                {processedResults.map((item, index) => {
                  if (!item.success) {
                    return (
                      <div key={index} className="glass-card rounded-xl p-5 border border-rose-500/30 text-xs flex items-center justify-between">
                        <div className="flex items-center space-x-2 text-rose-400 font-mono">
                          <AlertCircle className="w-4 h-4 shrink-0" />
                          <span>{item.file.name}</span>
                        </div>
                        <span className="px-2 py-0.5 rounded bg-rose-500/20 text-rose-300 font-medium">Failed: {item.error}</span>
                      </div>
                    );
                  }

                  const data = item.data;
                  const diffs = data.sampleModifications || [];
                  const isDiffOpen = openDiffIndex === index;

                  return (
                    <div key={index} className="glass-card rounded-xl p-5 border border-slate-800 space-y-4">
                      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4">
                        <div className="flex items-center space-x-3 truncate">
                          <div className="w-10 h-10 rounded-xl bg-emerald-500/20 text-emerald-400 flex items-center justify-center shrink-0">
                            <FileSpreadsheet className="w-5 h-5" />
                          </div>
                          <div className="truncate">
                            <div className="font-bold text-sm text-slate-100 font-mono truncate">{data.cleanedFilename}</div>
                            <div className="flex items-center space-x-2 text-xs text-slate-400 font-mono mt-0.5">
                              <span>{formatBytes(data.cleanedSize)}</span>
                              <span>•</span>
                              <span>{data.stats.sheetsProcessed} sheet(s)</span>
                              <span>•</span>
                              <span>{data.stats.rowsProcessed.toLocaleString()} rows</span>
                              <span>•</span>
                              <span className="text-emerald-400 font-semibold">{data.stats.modifiedCells.toLocaleString()} modified</span>
                            </div>
                          </div>
                        </div>

                        <div className="flex items-center space-x-2 shrink-0">
                          {diffs.length > 0 && (
                            <button
                              type="button"
                              onClick={() => setOpenDiffIndex(isDiffOpen ? null : index)}
                              className="px-3 py-1.5 rounded-lg text-xs font-medium text-slate-300 hover:text-white bg-slate-800 hover:bg-slate-700 border border-slate-700 transition flex items-center space-x-1"
                            >
                              <Eye className="w-3.5 h-3.5 text-emerald-400" />
                              <span>Sample Changes ({diffs.length})</span>
                            </button>
                          )}
                          <a
                            href={data.downloadUrl}
                            download={data.cleanedFilename}
                            className="px-3.5 py-1.5 rounded-lg text-xs font-semibold text-emerald-300 hover:text-white bg-emerald-500/10 hover:bg-emerald-500/20 border border-emerald-500/30 transition flex items-center space-x-1"
                          >
                            <Download className="w-3.5 h-3.5" />
                            <span>Download</span>
                          </a>
                        </div>
                      </div>

                      {/* Expandable Diffs */}
                      {isDiffOpen && diffs.length > 0 && (
                        <div className="pt-3 border-t border-slate-800/80">
                          <div className="text-[11px] font-bold text-slate-400 uppercase tracking-wider mb-2">Sample Cleaned Cells:</div>
                          <div className="max-h-48 overflow-y-auto space-y-1 pr-1">
                            {diffs.slice(0, 15).map((d, dIdx) => (
                              <div key={dIdx} className="p-2 rounded bg-slate-900/80 border border-slate-800 text-[11px] font-mono flex items-center justify-between gap-3">
                                <span className="text-emerald-400 font-semibold shrink-0">
                                  {d.sheet}!{d.cell}
                                </span>
                                <span className="diff-removed truncate text-right">{d.before}</span>
                                <span className="text-slate-500 shrink-0">➔</span>
                                <span className="diff-added truncate">{d.after}</span>
                              </div>
                            ))}
                          </div>
                        </div>
                      )}
                    </div>
                  );
                })}
              </div>
            </div>
          </div>
        )}

        {/* History Modal */}
        {showHistory && (
          <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-950/80 backdrop-blur-sm">
            <div className="glass-card rounded-2xl max-w-lg w-full p-6 shadow-2xl space-y-4 border border-slate-700">
              <div className="flex items-center justify-between">
                <h3 className="text-base font-bold text-white flex items-center space-x-2">
                  <History className="w-4 h-4 text-emerald-400" />
                  <span>Cleaned Files in Output Folder</span>
                </h3>
                <button onClick={() => setShowHistory(false)} className="text-slate-400 hover:text-white p-1 rounded-lg">
                  <X className="w-5 h-5" />
                </button>
              </div>
              <p className="text-xs text-slate-400">All saved files on disk in output folder:</p>

              <div className="max-h-64 overflow-y-auto divide-y divide-slate-800 space-y-1">
                {loadingHistory ? (
                  <div className="py-4 text-center text-xs text-slate-500">Loading files...</div>
                ) : historyFiles.length > 0 ? (
                  historyFiles.map((f, fIdx) => (
                    <div key={fIdx} className="py-2.5 flex items-center justify-between text-xs hover:bg-slate-800/50 px-2 rounded-lg transition">
                      <div className="truncate mr-3">
                        <div className="font-medium text-slate-200 truncate font-mono">{f.filename}</div>
                        <div className="text-[11px] text-slate-500 font-mono">
                          {formatBytes(f.size)} • {new Date(f.mtime).toLocaleString()}
                        </div>
                      </div>
                      <a
                        href={`/api/download/${encodeURIComponent(f.filename)}`}
                        download={f.filename}
                        className="px-2.5 py-1 bg-emerald-500/10 text-emerald-400 hover:bg-emerald-500/20 rounded-md border border-emerald-500/30 font-medium shrink-0 flex items-center space-x-1"
                      >
                        <Download className="w-3.5 h-3.5" />
                        <span>Download</span>
                      </a>
                    </div>
                  ))
                ) : (
                  <div className="py-4 text-center text-xs text-slate-500">No output files found.</div>
                )}
              </div>

              <div className="pt-2 text-right">
                <button
                  onClick={() => setShowHistory(false)}
                  className="px-4 py-2 bg-slate-800 hover:bg-slate-700 text-xs font-semibold text-slate-200 rounded-lg"
                >
                  Close
                </button>
              </div>
            </div>
          </div>
        )}
      </main>

      {/* Footer */}
      <footer className="border-t border-slate-800/80 bg-slate-900/40 py-6 text-center text-xs text-slate-500 relative z-10">
        <div className="max-w-6xl mx-auto px-4 flex flex-col sm:flex-row items-center justify-between gap-3">
          <div>Excel Cleaner Studio • Next.js Frontend + Fastify Backend</div>
          <div className="text-slate-400 flex items-center space-x-4 font-mono text-[11px]">
            <span>Vercel Deployable</span>
          </div>
        </div>
      </footer>

      {/* Toast Notification */}
      {toast && (
        <div
          className={`fixed bottom-6 right-6 z-50 max-w-sm rounded-xl p-4 shadow-2xl flex items-center space-x-3 text-xs transition-all duration-300 ${
            toast.isError
              ? 'bg-rose-950 border border-rose-500/50 text-rose-200'
              : 'bg-emerald-950 border border-emerald-500/50 text-emerald-200'
          }`}
        >
          {toast.isError ? (
            <AlertCircle className="w-5 h-5 text-rose-400 shrink-0" />
          ) : (
            <CheckCircle2 className="w-5 h-5 text-emerald-400 shrink-0" />
          )}
          <span className="font-medium">{toast.message}</span>
        </div>
      )}
    </div>
  );
}
