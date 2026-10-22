(function() {
  if (typeof window.__echoPowerPointBridgeDispose === 'function') {
    try {
      window.__echoPowerPointBridgeDispose();
    } catch {}
  }
  window.__echoPowerPointBridgeInstalled = true;

  function installHostFixes() {
    try {
      let style = document.getElementById('echo-ppt-host-fix');
      if (!style) {
        style = document.createElement('style');
        style.id = 'echo-ppt-host-fix';
        (document.head || document.documentElement).appendChild(style);
      }
      if (style.dataset.applied === '1') return;
      style.dataset.applied = '1';
      style.textContent = `
        /* 仅在工坊全屏歌词模式下穿透宿主原生标题栏及品牌层，防止吞噬中间搜索框与右上角按钮点击 */
        body:has(.workshop-ui-runtime) header.app-titlebar,
        body:has(.workshop-ui-runtime--lyrics) header.app-titlebar,
        body:has(.workshop-lyrics-runtime-target:not(:empty)) header.app-titlebar,
        body:has(iframe[src*="echo.powerpoint-theme"]) header.app-titlebar,
        body:has(iframe.workshop-ui-runtime__frame) header.app-titlebar,
        body:has(.workshop-plugin-runtimes iframe) header.app-titlebar {
          pointer-events: none !important;
          -webkit-app-region: no-drag !important;
        }
        body:has(.workshop-ui-runtime) .app-titlebar-brand,
        body:has(.workshop-ui-runtime--lyrics) .app-titlebar-brand,
        body:has(.workshop-lyrics-runtime-target:not(:empty)) .app-titlebar-brand,
        body:has(iframe[src*="echo.powerpoint-theme"]) .app-titlebar-brand,
        body:has(iframe.workshop-ui-runtime__frame) .app-titlebar-brand,
        body:has(.workshop-plugin-runtimes iframe) .app-titlebar-brand {
          pointer-events: none !important;
          -webkit-app-region: no-drag !important;
        }
        /* 隐藏宿主原生居中快速切换下拉控件与异常报告气泡 */
        .lyrics-style-quick-switch,
        .titlebar-report-notice {
          display: none !important;
          pointer-events: none !important;
        }
        /* 保留宿主原生窗口操作（最小化、最大化、关闭、系统托盘、返回） */
        .app-titlebar-actions,
        .window-controls,
        .app-titlebar-actions button,
        .window-controls button,
        .workshop-lyrics-runtime-back {
          pointer-events: auto !important;
          -webkit-app-region: no-drag !important;
        }
        /* 提升工坊全屏歌词运行时层级，确保PPT主题各控制按钮正常响应 */
        .workshop-ui-runtime,
        .workshop-lyrics-runtime-target,
        section.workshop-ui-runtime {
          z-index: 60 !important;
        }
      `;
    } catch (e) {}
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', installHostFixes);
  } else {
    installHostFixes();
  }
  const hostFixTimer = setInterval(installHostFixes, 1000);

  const cache = new Map();
  let lastBroadcastTrackId = null;
  let lastBroadcastCoverUrl = null;
  let lastQueueSignature = '';
  let lastPlayTrackTime = 0;
  let lastKnownQueueItems = [];

  async function resolveCoverDataUrl(targetUrl, isThumb = false) {
    if (!targetUrl) {
      const status = await window.echo?.audio?.getStatus?.();
      targetUrl = status?.currentTrackCoverUrl;
    }
    if (!targetUrl) {
      const img = document.querySelector('.player-cover-visual img') ||
                  document.querySelector('.player-cover img, [class*="player-cover"] img') ||
                  document.querySelector('.player-bar img, .player img, [class*="cover"] img');
      if (img && img.src) targetUrl = img.src;
    }
    if (!targetUrl) return null;

    // High resolution upgrade: if requesting full cover, upgrade to large (768x768); for thumb, upgrade to album (320x320)
    if (typeof targetUrl === 'string') {
      if (!isThumb) {
        targetUrl = targetUrl.replace('/thumb/', '/large/').replace('/album/', '/large/');
      } else if (targetUrl.includes('/thumb/')) {
        targetUrl = targetUrl.replace('/thumb/', '/album/');
      }
    }

    const cacheKey = (isThumb ? 'thumb:' : 'full:') + targetUrl;
    if (cache.has(cacheKey)) return cache.get(cacheKey);

    // Direct Image loading (handles echo-cover:// protocol seamlessly in renderer)
    try {
      if (targetUrl.startsWith('echo-cover:') || targetUrl.startsWith('http://') || targetUrl.startsWith('https://') || targetUrl.startsWith('data:')) {
        const offImg = new Image();
        offImg.crossOrigin = 'anonymous';
        const loaded = await new Promise((res) => {
          offImg.onload = () => res(true);
          offImg.onerror = () => res(false);
          offImg.src = targetUrl;
        });
        if (loaded && offImg.naturalWidth > 0) {
          const canvas = document.createElement('canvas');
          if (isThumb) {
            canvas.width = 320;
            canvas.height = 180;
            const ctx = canvas.getContext('2d');
            const scale = Math.max(320 / offImg.naturalWidth, 180 / offImg.naturalHeight);
            const w = offImg.naturalWidth * scale;
            const h = offImg.naturalHeight * scale;
            const x = (320 - w) / 2;
            const y = (180 - h) / 2;
            ctx.drawImage(offImg, x, y, w, h);
            const dataUrl = canvas.toDataURL('image/jpeg', 0.9);
            if (dataUrl) {
              cache.set(cacheKey, dataUrl);
              return dataUrl;
            }
          } else {
            canvas.width = offImg.naturalWidth;
            canvas.height = offImg.naturalHeight;
            const ctx = canvas.getContext('2d');
            ctx.drawImage(offImg, 0, 0);
            const dataUrl = canvas.toDataURL('image/jpeg', 0.95);
            if (dataUrl) {
              cache.set(cacheKey, dataUrl);
              return dataUrl;
            }
          }
        }
      }
    } catch (errImgLoad) {}

    try {
      const resp = await fetch(targetUrl);
      if (resp.ok) {
        const blob = await resp.blob();
        if (isThumb) {
          try {
            const bitmap = await createImageBitmap(blob);
            const canvas = document.createElement('canvas');
            canvas.width = 320;
            canvas.height = 180;
            const ctx = canvas.getContext('2d');
            const scale = Math.max(320 / bitmap.width, 180 / bitmap.height);
            const w = bitmap.width * scale;
            const h = bitmap.height * scale;
            const x = (320 - w) / 2;
            const y = (180 - h) / 2;
            ctx.drawImage(bitmap, x, y, w, h);
            const thumbDataUrl = canvas.toDataURL('image/jpeg', 0.9);
            if (thumbDataUrl) {
              cache.set(cacheKey, thumbDataUrl);
              if (cache.size > 200) {
                const first = cache.keys().next().value;
                cache.delete(first);
              }
              return thumbDataUrl;
            }
          } catch (errBitmap) {}
        }
        const reader = new FileReader();
        const dataUrl = await new Promise((resolve) => {
          reader.onloadend = () => resolve(reader.result);
          reader.readAsDataURL(blob);
        });
        if (dataUrl) {
          cache.set(cacheKey, dataUrl);
          if (cache.size > 200) {
            const first = cache.keys().next().value;
            cache.delete(first);
          }
          return dataUrl;
        }
      }
    } catch {}

    try {
      const img = document.querySelector('.player-cover-visual img') ||
                  document.querySelector('.player-cover img, [class*="player-cover"] img') ||
                  document.querySelector('.player-bar img, .player img, [class*="cover"] img');
      if (img && img.naturalWidth > 0) {
        const canvas = document.createElement('canvas');
        if (isThumb) {
          canvas.width = 320;
          canvas.height = 180;
          const ctx = canvas.getContext('2d');
          const scale = Math.max(320 / img.naturalWidth, 180 / img.naturalHeight);
          const w = img.naturalWidth * scale;
          const h = img.naturalHeight * scale;
          const x = (320 - w) / 2;
          const y = (180 - h) / 2;
          ctx.drawImage(img, x, y, w, h);
          const dataUrl = canvas.toDataURL('image/jpeg', 0.9);
          if (dataUrl) {
            cache.set(cacheKey, dataUrl);
            return dataUrl;
          }
        } else {
          canvas.width = img.naturalWidth;
          canvas.height = img.naturalHeight;
          const ctx = canvas.getContext('2d');
          ctx.drawImage(img, 0, 0);
          const dataUrl = canvas.toDataURL('image/jpeg', 0.95);
          if (dataUrl) {
            cache.set(cacheKey, dataUrl);
            return dataUrl;
          }
        }
      }
    } catch {}

    return null;
  }

  function parseLrc(lrcText) {
    if (!lrcText || typeof lrcText !== 'string') return [];
    const lines = [];
    const timeRegex = /\[(\d{1,2}):(\d{2})(?:\.(\d{1,3}))?\]/g;
    const rows = lrcText.split(/\r?\n/);
    for (const row of rows) {
      const trimmed = row.trim();
      if (!trimmed) continue;
      let match;
      const timestamps = [];
      timeRegex.lastIndex = 0;
      while ((match = timeRegex.exec(trimmed)) !== null) {
        const min = parseInt(match[1], 10);
        const sec = parseInt(match[2], 10);
        let ms = 0;
        if (match[3]) {
          ms = parseInt(match[3].padEnd(3, '0').slice(0, 3), 10);
        }
        timestamps.push(min * 60000 + sec * 1000 + ms);
      }
      const cleanText = trimmed.replace(timeRegex, '').trim();
      if (timestamps.length > 0) {
        for (const timeMs of timestamps) {
          lines.push({ timeMs, text: cleanText });
        }
      } else if (cleanText) {
        lines.push({ timeMs: 0, text: cleanText });
      }
    }
    lines.sort((a, b) => a.timeMs - b.timeMs);
    return lines;
  }

  async function resolveLyrics(trackId) {
    if (!trackId) {
      const status = await window.echo?.audio?.getStatus?.();
      trackId = status?.currentTrackId;
    }
    if (!trackId || typeof trackId !== 'string') return null;

    // 1. Try host lyrics manager (handles local tracks with embedded/synced lyrics)
    try {
      const res = await window.echo?.lyrics?.getForTrack?.(trackId);
      if (res) {
        let lines = [];
        if (Array.isArray(res.lines) && res.lines.length > 0) {
          lines = res.lines.map((l) => ({
            timeMs: Number(l.timeMs) || 0,
            text: String(l.text || ''),
            translation: l.translation ? String(l.translation) : null,
            romanization: l.romanization ? String(l.romanization) : null
          }));
        } else if (res.syncedText) {
          lines = parseLrc(res.syncedText);
        } else if (res.plainText) {
          lines = parseLrc(res.plainText);
        }

        const hasTiming = lines.some((l) => l.timeMs > 0);
        const kind = res.kind === 'instrumental' ? 'instrumental' : (hasTiming ? 'synced' : (lines.length > 0 ? 'plain' : 'empty'));

        return {
          kind,
          title: res.title || '',
          artist: res.artist || '',
          album: res.album || null,
          durationSeconds: res.durationSeconds || null,
          offsetMs: res.offsetMs || 0,
          provider: res.provider || 'local',
          lines,
          plainText: res.plainText || null,
          syncedText: res.syncedText || null
        };
      }
    } catch {}

    // 2. Try streaming lyrics provider
    if (trackId.startsWith('streaming:')) {
      const parts = trackId.split(':');
      if (parts.length >= 3) {
        const provider = parts[1];
        const providerTrackId = parts.slice(2).join(':');

        try {
          const res = await window.echo?.streaming?.getLyrics?.({
            provider,
            providerTrackId,
            id: trackId
          });
          if (res) {
            let lines = [];
            if (Array.isArray(res.lines) && res.lines.length > 0) {
              lines = res.lines.map((l) => ({
                timeMs: Number(l.timeMs) || 0,
                text: String(l.text || ''),
                translation: l.translation ? String(l.translation) : null,
                romanization: l.romanization ? String(l.romanization) : null
              }));
            } else if (res.syncedLyrics) {
              lines = parseLrc(res.syncedLyrics);
            } else if (res.plainLyrics) {
              lines = parseLrc(res.plainLyrics);
            }

            const hasTiming = lines.some((l) => l.timeMs > 0);
            const kind = res.instrumental ? 'instrumental' : (hasTiming ? 'synced' : (lines.length > 0 ? 'plain' : 'empty'));

            return {
              kind,
              title: res.title || '',
              artist: res.artist || '',
              album: res.album || null,
              durationSeconds: null,
              offsetMs: 0,
              provider: res.sourceLabel || provider,
              lines,
              plainText: res.plainLyrics || null,
              syncedText: res.syncedLyrics || null
            };
          }
        } catch {}
      }
    }

    // 3. Fallback: query library track metadata for embedded lyrics
    try {
      const track = await window.echo?.library?.getTrack?.(trackId);
      if (track) {
        const rawLrc = track.syncedLyrics || track.lyrics || track.plainLyrics || track.embeddedLyrics;
        if (rawLrc) {
          const lines = parseLrc(rawLrc);
          const hasTiming = lines.some((l) => l.timeMs > 0);
          return {
            kind: hasTiming ? 'synced' : 'plain',
            title: track.title || '',
            artist: track.artist || '',
            album: track.album || null,
            durationSeconds: track.duration || null,
            offsetMs: 0,
            provider: 'embedded',
            lines,
            plainText: rawLrc,
            syncedText: hasTiming ? rawLrc : null
          };
        }
      }
    } catch {}

    return null;
  }

  function broadcastCover(trackId, dataUrl) {
    if (!dataUrl) return;
    const iframes = document.querySelectorAll('iframe.workshop-ui-runtime__frame, iframe');
    for (const frame of iframes) {
      try {
        frame.contentWindow?.postMessage({
          type: 'echo:powerpoint:coverData',
          trackId,
          coverDataUrl: dataUrl
        }, '*');
      } catch {}
    }
  }

  function broadcastLyrics(trackId, lyrics) {
    if (!lyrics) return;
    const iframes = document.querySelectorAll('iframe.workshop-ui-runtime__frame, iframe');
    for (const frame of iframes) {
      try {
        frame.contentWindow?.postMessage({
          type: 'echo:powerpoint:lyricsData',
          trackId,
          lyrics
        }, '*');
      } catch {}
    }
  }

  function broadcastQueueCover(trackId, queueId, rawUrl, thumbData) {
    if (!thumbData) return;
    const iframes = document.querySelectorAll('iframe.workshop-ui-runtime__frame, iframe');
    for (const frame of iframes) {
      try {
        frame.contentWindow?.postMessage({
          type: 'echo:powerpoint:queueCoverItem',
          trackId,
          queueId,
          rawUrl,
          coverDataUrl: thumbData
        }, '*');
      } catch {}
    }
  }

  async function resolveQueueCoversInBackground(items) {
    if (!items || !items.length) return;
    const toResolve = items.filter(it => it.rawCoverUrl && (it.rawCoverUrl.startsWith('http://') || it.rawCoverUrl.startsWith('https://') || it.rawCoverUrl.startsWith('echo-cover:')) && !cache.has('thumb:' + it.rawCoverUrl));
    if (!toResolve.length) return;

    for (let i = 0; i < toResolve.length; i += 3) {
      const batch = toResolve.slice(i, i + 3);
      await Promise.all(batch.map(async (item) => {
        try {
          const thumbData = await resolveCoverDataUrl(item.rawCoverUrl, true);
          if (thumbData) {
            broadcastQueueCover(item.id, item.queueId, item.rawCoverUrl, thumbData);
          }
        } catch {}
      }));
    }
  }

  async function fetchAllLibraryTracks(maxPages = 3, pageSize = 150) {
    const collected = [];
    const lib = window.echo?.library;
    if (!lib || typeof lib.getTracks !== 'function') return collected;
    for (let page = 1; page <= maxPages; page++) {
      let res = null;
      try {
        res = await lib.getTracks({ page, pageSize });
      } catch {
        if (page === 1) {
          try { res = await lib.getTracks(); } catch {}
        }
      }
      const list = Array.isArray(res?.items) ? res.items : (Array.isArray(res) ? res : []);
      if (!list.length) break;
      collected.push(...list);
      const total = Number(res?.total || 0);
      if ((total && collected.length >= total) || list.length < pageSize) break;
    }
    return collected;
  }

  function unwrapTrackObject(rawItem, idx = 0) {
    const tr = rawItem?.track || rawItem || {};
    const rawCover = tr.coverUrl || rawItem?.coverUrl || tr.coverThumb || rawItem?.coverThumb || '';
    const normCover = typeof rawCover === 'string' ? rawCover.replace('/thumb/', '/album/') : '';
    return {
      id: tr.id || rawItem?.trackId || rawItem?.id || ('track-' + idx),
      queueId: rawItem?.queueId || null,
      title: tr.title || rawItem?.title || ('Track ' + (idx + 1)),
      artist: tr.artist || rawItem?.artist || '—',
      album: tr.album || rawItem?.album || '—',
      durationSeconds: Number(tr.durationSeconds || tr.duration || rawItem?.durationSeconds || rawItem?.duration || 0),
      coverThumb: normCover,
      coverUrl: normCover,
      rawTrack: tr
    };
  }

  function synthesizeAlbumsFromTrackList(tracks) {
    const map = new Map();
    for (const raw of (tracks || [])) {
      const t = unwrapTrackObject(raw);
      const albumTitle = (t.album && t.album !== '—') ? t.album : '未分类专辑';
      const key = albumTitle.toLowerCase();
      if (!map.has(key)) {
        map.set(key, {
          id: 'synth-album:' + encodeURIComponent(albumTitle),
          title: albumTitle,
          artist: (t.artist && t.artist !== '—') ? t.artist : '未知艺术家',
          year: '',
          itemCount: 0,
          durationSeconds: 0,
          coverThumb: t.coverThumb || ''
        });
      }
      const entry = map.get(key);
      entry.itemCount += 1;
      entry.durationSeconds += Number(t.durationSeconds || 0);
      if (!entry.coverThumb && t.coverThumb) entry.coverThumb = t.coverThumb;
    }
    return Array.from(map.values());
  }

  async function broadcastQueue(force = false) {
    let items = [];
    try {
      const repo = window.echo?.playback?.queueRepository;
      if (repo && typeof repo.read === 'function') {
        let offset = 0;
        const limit = 200;
        while (true) {
          const page = await repo.read({ operation: 'page', limit, offset });
          if (!page?.items?.length) break;
          items.push(...page.items);
          offset += page.items.length;
          if (offset >= (page.total || 0) || page.items.length < limit) break;
        }
      }
    } catch {}

    if (!items.length) {
      try {
        const sessionQ = await window.echo?.playback?.getQueueSession?.();
        if (sessionQ?.items?.length) items = sessionQ.items;
      } catch {}
    }
    if (!items.length) {
      try {
        const miniQ = await window.echo?.playback?.miniPlayerQueueRead?.();
        if (miniQ?.items?.length) items = miniQ.items;
      } catch {}
    }
    if (!items.length) {
      const player = window.__echoExternalPlayer;
      const q = player?.queue ? (typeof player.queue === 'function' ? player.queue() : player.queue) : null;
      if (Array.isArray(q?.items) && q.items.length) {
        items = q.items;
      }
    }
    if (!items.length) {
      try {
        const status = await window.echo?.audio?.getStatus?.();
        if (status?.currentTrackId || status?.currentTrackTitle) {
          items = [{
            id: status.currentTrackId || 'current',
            queueId: 'current',
            title: status.currentTrackTitle || '当前播放曲目',
            artist: status.currentTrackArtist || '—',
            album: status.currentTrackAlbum || '—',
            durationSeconds: status.durationSeconds || 0,
            coverUrl: status.currentTrackCoverUrl || ''
          }];
        }
      } catch {}
    }

    const mappedItems = items.map((it, idx) => {
      const track = it.track || it;
      let rawCover = track.coverUrl || it.coverUrl || track.coverThumb || it.coverThumb || '';
      if (typeof rawCover === 'string' && rawCover.includes('echo-cover://thumb/')) {
        rawCover = rawCover.replace('echo-cover://thumb/', 'echo-cover://album/');
      }
      const cachedThumb = rawCover ? cache.get('thumb:' + rawCover) : null;
      return {
        queueId: it.queueId || ('q-' + (track.id || it.id || idx)),
        id: track.id || it.id,
        title: track.title || it.title || ('Slide ' + (idx + 1)),
        artist: track.artist || it.artist || '—',
        album: track.album || it.album || '—',
        durationSeconds: track.durationSeconds || track.duration || it.durationSeconds || it.duration || 0,
        rawCoverUrl: rawCover,
        coverThumb: cachedThumb || (rawCover.startsWith('echo-cover:') || rawCover.startsWith('data:') ? rawCover : ''),
        coverUrl: cachedThumb || (rawCover.startsWith('echo-cover:') || rawCover.startsWith('data:') ? rawCover : '')
      };
    });

    if (mappedItems.length > 0) {
      lastKnownQueueItems = mappedItems;
    }

    const iframes = document.querySelectorAll('iframe.workshop-ui-runtime__frame, iframe[src*="echo.powerpoint-theme"], iframe[src*="echo-workshop://ui"], iframe');
    if (!iframes.length) {
      return;
    }

    const signature = mappedItems.map(it => (it.queueId || '') + ':' + (it.id || '')).join('|');
    if (!force && signature === lastQueueSignature) {
      return;
    }
    lastQueueSignature = signature;

    for (const frame of iframes) {
      try {
        frame.contentWindow?.postMessage({
          type: 'echo:powerpoint:queueData',
          items: mappedItems
        }, '*');
      } catch {}
    }

    resolveQueueCoversInBackground(mappedItems);
  }

  async function broadcastAudioStatus(status) {
    if (!status) {
      try {
        status = await window.echo?.audio?.getStatus?.();
      } catch {}
    }
    if (!status) return;
    const iframes = document.querySelectorAll('iframe.workshop-ui-runtime__frame, iframe');
    for (const frame of iframes) {
      try {
        frame.contentWindow?.postMessage({
          type: 'echo:powerpoint:audioStatus',
          status
        }, '*');
      } catch {}
    }
  }

  function postToAllFrames(payload) {
    const iframes = document.querySelectorAll('iframe.workshop-ui-runtime__frame, iframe[src*="echo.powerpoint-theme"], iframe[src*="echo-workshop://ui"], iframe');
    for (const frame of iframes) {
      try {
        frame.contentWindow?.postMessage(payload, '*');
      } catch {}
    }
  }

  // --- Persistent User Preset Storage Bridge ---
  const PRESET_STORAGE_KEY = 'echo_powerpoint_user_preset';
  const FONT_STORAGE_KEY = 'echo_powerpoint_lyrics_font';

  function getStoredPreset() {
    try {
      const raw = localStorage.getItem(PRESET_STORAGE_KEY);
      return raw ? JSON.parse(raw) : null;
    } catch {
      return null;
    }
  }

  function getStoredFont() {
    try {
      const raw = localStorage.getItem(FONT_STORAGE_KEY);
      return raw ? JSON.parse(raw) : null;
    } catch {
      return null;
    }
  }

  function saveStoredPreset(preset) {
    try {
      if (preset && typeof preset === 'object') {
        localStorage.setItem(PRESET_STORAGE_KEY, JSON.stringify(preset));
      }
    } catch {}
  }

  function saveStoredFont(font) {
    try {
      if (font && typeof font === 'object') {
        localStorage.setItem(FONT_STORAGE_KEY, JSON.stringify(font));
      }
    } catch {}
  }

  function broadcastPreset(targetWindow) {
    const payload = {
      type: 'echo:powerpoint:presetData',
      preset: getStoredPreset(),
      fontSettings: getStoredFont()
    };
    postToAllFrames(payload);
    if (targetWindow && typeof targetWindow.postMessage === 'function') {
      try { targetWindow.postMessage(payload, '*'); } catch {}
    }
  }

  // --- IPC Window Message Listener ---
  const onWindowMessage = async (e) => {
    if (!e.data || typeof e.data !== 'object') return;
    const type = e.data.type || '';

    // Preset Memory & Persistence
    if (type === 'echo:powerpoint:getPreset' || type === 'echo:workshop-ui:ready') {
      broadcastPreset(e.source || null);
    }

    if (type === 'echo:powerpoint:savePreset') {
      if (e.data.preset) {
        saveStoredPreset(e.data.preset);
      }
    }

    if (type === 'echo:powerpoint:saveLyricsFont') {
      if (e.data.fontSettings !== undefined) {
        saveStoredFont(e.data.fontSettings);
      }
    }

    if (type === 'echo:powerpoint:getAudioStatus' || type === 'echo:audition:getAudioStatus') {
      try {
        broadcastPreset(e.source || null);
        const status = await window.echo?.audio?.getStatus?.();
        if (status) broadcastAudioStatus(status);
      } catch {}
    }

    if (type === 'echo:powerpoint:getCover' || type === 'echo:audition:getCover') {
      const trackId = e.data.trackId || (await window.echo?.audio?.getStatus?.())?.currentTrackId;
      const url = e.data.url;
      const dataUrl = await resolveCoverDataUrl(url);
      if (dataUrl) {
        broadcastCover(trackId, dataUrl);
      }
    }

    if (type === 'echo:powerpoint:getLyrics' || type === 'echo:audition:getLyrics') {
      const trackId = e.data.trackId || (await window.echo?.audio?.getStatus?.())?.currentTrackId;
      const lyrics = await resolveLyrics(trackId);
      if (lyrics) {
        broadcastLyrics(trackId, lyrics);
      }
    }

    if (type === 'echo:powerpoint:getQueue' || type === 'echo:audition:getQueue') {
      broadcastQueue(true);
    }

    if (type === 'echo:powerpoint:copyText' && e.data?.text) {
      try {
        if (navigator.clipboard?.writeText) {
          navigator.clipboard.writeText(e.data.text).catch(() => {});
        }
      } catch {}
    }

    // Read installed system fonts
    if (type === 'echo:powerpoint:getSystemFonts') {
      let fontNames = [];
      try {
        if (typeof window.queryLocalFonts === 'function') {
          const fonts = await window.queryLocalFonts();
          const set = new Set();
          for (const f of fonts) {
            if (f.family) set.add(f.family);
          }
          fontNames = Array.from(set).sort();
        }
      } catch {}
      postToAllFrames({
        type: 'echo:powerpoint:systemFonts',
        fonts: fontNames
      });
    }

    // Full Library Tracks for Global Search
    if (type === 'echo:powerpoint:getFullLibrary') {
      let mapped = [];
      try {
        let tracks = await fetchAllLibraryTracks(3, 150);
        if (!tracks.length && lastKnownQueueItems.length) {
          tracks = lastKnownQueueItems;
        }
        mapped = tracks.map((t, idx) => {
          const u = unwrapTrackObject(t, idx);
          return {
            id: u.id,
            title: u.title,
            artist: u.artist,
            album: u.album,
            durationSeconds: u.durationSeconds,
            coverThumb: u.coverThumb,
            coverUrl: u.coverUrl
          };
        });
      } catch {}
      postToAllFrames({
        type: 'echo:powerpoint:fullLibrary',
        tracks: mapped
      });
    }

    // Read Playlists, Albums, and Library Folders (each wrapped in isolated try/catch)
    if (type === 'echo:powerpoint:getPlaylists') {
      let rawPlaylists = [];
      let rawFolders = [];
      let rawAlbums = [];

      try {
        let plRes = null;
        try {
          plRes = await window.echo?.library?.getPlaylists?.({ page: 1, pageSize: 100 });
        } catch {
          plRes = await window.echo?.library?.getPlaylists?.();
        }
        rawPlaylists = Array.isArray(plRes?.items) ? plRes.items : (Array.isArray(plRes) ? plRes : []);
      } catch {}

      try {
        const fRes = await window.echo?.library?.getFolders?.();
        rawFolders = Array.isArray(fRes?.items) ? fRes.items : (Array.isArray(fRes) ? fRes : []);
      } catch {}

      try {
        let albRes = null;
        try {
          albRes = await window.echo?.library?.getAlbums?.({ page: 1, pageSize: 100 });
        } catch {
          albRes = await window.echo?.library?.getAlbums?.();
        }
        rawAlbums = Array.isArray(albRes?.items) ? albRes.items : (Array.isArray(albRes) ? albRes : []);
      } catch {}

      const playlists = rawPlaylists.map(p => ({
        id: p.id,
        name: p.name || p.title || '未命名歌单',
        itemCount: Number(p.itemCount ?? p.trackCount ?? (Array.isArray(p.items) ? p.items.length : 0)),
        coverThumb: p.coverThumb ? String(p.coverThumb).replace('/thumb/', '/album/') : ''
      }));

      const folders = rawFolders.map(f => ({
        id: f.id || f.path || 'folder',
        name: f.name || (f.path ? String(f.path).split(/[\\/]/).pop() : '曲库文件夹'),
        path: f.path || ''
      }));

      let albums = rawAlbums.map(a => ({
        id: a.id || ('album:' + (a.title || '')),
        title: a.title || a.name || '未知专辑',
        artist: a.albumArtist || a.artist || '未知艺术家',
        year: a.year || '',
        itemCount: Number(a.trackCount ?? a.itemCount ?? 0),
        durationSeconds: Number(a.duration ?? a.durationSeconds ?? 0),
        coverThumb: a.coverThumb ? String(a.coverThumb).replace('/thumb/', '/album/') : ''
      }));

      if (!albums.length) {
        try {
          let fallbackTracks = await fetchAllLibraryTracks(2, 150);
          if (!fallbackTracks.length && lastKnownQueueItems.length) {
            fallbackTracks = lastKnownQueueItems;
          }
          if (fallbackTracks.length) {
            albums = synthesizeAlbumsFromTrackList(fallbackTracks);
          }
        } catch {}
      }

      postToAllFrames({
        type: 'echo:powerpoint:playlists',
        playlists,
        folders,
        albums
      });
    }

    // Read specific playlist items
    if (type === 'echo:powerpoint:getPlaylistItems') {
      const playlistId = e.data.playlistId;
      let items = [];
      try {
        let pItems = null;
        try {
          pItems = await window.echo?.library?.getPlaylistItems?.(playlistId, { page: 1, pageSize: 200 });
        } catch {
          pItems = await window.echo?.library?.getPlaylistItems?.(playlistId);
        }
        const raw = Array.isArray(pItems?.items) ? pItems.items : (Array.isArray(pItems) ? pItems : []);
        items = raw.map((t, idx) => {
          const u = unwrapTrackObject(t, idx);
          return {
            id: u.id,
            title: u.title,
            artist: u.artist,
            album: u.album,
            durationSeconds: u.durationSeconds,
            coverThumb: u.coverThumb
          };
        });
      } catch {}
      postToAllFrames({
        type: 'echo:powerpoint:playlistItems',
        playlistId,
        items
      });
    }

    // Read specific album items
    if (type === 'echo:powerpoint:getAlbumTracks') {
      const albumId = e.data.albumId;
      let items = [];
      try {
        if (typeof albumId === 'string' && albumId.startsWith('synth-album:')) {
          const targetAlbum = decodeURIComponent(albumId.slice('synth-album:'.length)).toLowerCase();
          let allTracks = await fetchAllLibraryTracks(2, 150);
          if (!allTracks.length && lastKnownQueueItems.length) {
            allTracks = lastKnownQueueItems;
          }
          items = allTracks
            .map((t, idx) => unwrapTrackObject(t, idx))
            .filter(u => ((u.album && u.album !== '—') ? u.album : '未分类专辑').toLowerCase() === targetAlbum);
        } else {
          let aItems = null;
          try {
            aItems = await window.echo?.library?.getAlbumTracks?.(albumId, { page: 1, pageSize: 200 });
          } catch {
            aItems = await window.echo?.library?.getAlbumTracks?.(albumId);
          }
          const raw = Array.isArray(aItems?.items) ? aItems.items : (Array.isArray(aItems) ? aItems : []);
          items = raw.map((t, idx) => {
            const u = unwrapTrackObject(t, idx);
            return {
              id: u.id,
              title: u.title,
              artist: u.artist,
              album: u.album,
              durationSeconds: u.durationSeconds,
              coverThumb: u.coverThumb
            };
          });
        }
      } catch {}
      postToAllFrames({
        type: 'echo:powerpoint:albumItems',
        albumId,
        items
      });
    }

    // Import Playlist to Queue
    if (type === 'echo:powerpoint:importPlaylist') {
      try {
        const { playlistId, mode } = e.data;
        let pItems = null;
        try {
          pItems = await window.echo?.library?.getPlaylistItems?.(playlistId, { page: 1, pageSize: 200 });
        } catch {
          pItems = await window.echo?.library?.getPlaylistItems?.(playlistId);
        }
        const raw = Array.isArray(pItems?.items) ? pItems.items : (Array.isArray(pItems) ? pItems : []);
        const tracks = raw.map((t, idx) => unwrapTrackObject(t, idx).rawTrack).filter(Boolean);
        const p = window.__echoExternalPlayer;
        const q = typeof p?.queue === 'function' ? p.queue() : p?.queue;

        if (mode === 'append' && q?.appendTracksToQueue) {
          await q.appendTracksToQueue(tracks);
        } else if (q?.replaceQueue) {
          await q.replaceQueue(tracks);
          if (q.playTrack && tracks.length > 0 && tracks[0]?.id) {
            await q.playTrack(tracks[0].id);
          }
        }
        // Broadcast resolved items to iframe for fallback/direct ingestion
        const iframes = document.querySelectorAll('iframe.workshop-ui-runtime__frame, iframe');
        for (const frame of iframes) {
          try {
            frame.contentWindow?.postMessage({
              type: 'echo:powerpoint:importTracksResult',
              playlistId,
              mode,
              tracks: tracks.map((t, idx) => unwrapTrackObject(t, idx))
            }, '*');
          } catch {}
        }
        setTimeout(() => broadcastQueue(true), 120);
      } catch {}
    }

    // Import Album to Queue
    if (type === 'echo:powerpoint:importAlbum') {
      try {
        const { albumId, mode } = e.data;
        let tracks = [];
        if (typeof albumId === 'string' && albumId.startsWith('synth-album:')) {
          const targetAlbum = decodeURIComponent(albumId.slice('synth-album:'.length)).toLowerCase();
          let allTracks = await fetchAllLibraryTracks(2, 150);
          if (!allTracks.length && lastKnownQueueItems.length) {
            allTracks = lastKnownQueueItems;
          }
          tracks = allTracks
            .map((t, idx) => unwrapTrackObject(t, idx))
            .filter(u => ((u.album && u.album !== '—') ? u.album : '未分类专辑').toLowerCase() === targetAlbum)
            .map(u => u.rawTrack);
        } else {
          let aItems = null;
          try {
            aItems = await window.echo?.library?.getAlbumTracks?.(albumId, { page: 1, pageSize: 200 });
          } catch {
            aItems = await window.echo?.library?.getAlbumTracks?.(albumId);
          }
          const raw = Array.isArray(aItems?.items) ? aItems.items : (Array.isArray(aItems) ? aItems : []);
          tracks = raw.map((t, idx) => unwrapTrackObject(t, idx).rawTrack).filter(Boolean);
        }

        const p = window.__echoExternalPlayer;
        const q = typeof p?.queue === 'function' ? p.queue() : p?.queue;

        if (mode === 'append' && q?.appendTracksToQueue) {
          await q.appendTracksToQueue(tracks);
        } else if (q?.replaceQueue) {
          await q.replaceQueue(tracks);
          if (q.playTrack && tracks.length > 0 && tracks[0]?.id) {
            await q.playTrack(tracks[0].id);
          }
        }
        // Broadcast resolved items to iframe for fallback/direct ingestion
        const iframes = document.querySelectorAll('iframe.workshop-ui-runtime__frame, iframe');
        for (const frame of iframes) {
          try {
            frame.contentWindow?.postMessage({
              type: 'echo:powerpoint:importTracksResult',
              albumId,
              mode,
              tracks: tracks.map((t, idx) => unwrapTrackObject(t, idx))
            }, '*');
          } catch {}
        }
        setTimeout(() => broadcastQueue(true), 120);
      } catch {}
    }

    // Toggle Host Window Fullscreen
    if (type === 'echo:powerpoint:toggleFullscreen') {
      try {
        const force = e.data?.fullscreen;
        const fsBtn = document.querySelector('.window-control--fullscreen');
        if (fsBtn) {
          const isFs = fsBtn.getAttribute('data-fullscreen') === 'true' || fsBtn.getAttribute('aria-pressed') === 'true';
          if (force === undefined || force !== isFs) {
            fsBtn.click();
          }
        }
      } catch {}
    }

    // Global Library Search Bridge
    if (type === 'echo:powerpoint:searchTracks') {
      const query = (e.data?.query || '').trim();
      let tracks = [];
      if (query) {
        try {
          const res = await window.echo?.library?.getTracks?.({ search: query, pageSize: 40 });
          tracks = Array.isArray(res?.items) ? res.items : (Array.isArray(res?.tracks) ? res.tracks : []);
        } catch {}
      }
      const iframes = document.querySelectorAll('iframe.workshop-ui-runtime__frame, iframe');
      for (const frame of iframes) {
        try {
          frame.contentWindow?.postMessage({
            type: 'echo:powerpoint:searchResults',
            query,
            tracks: tracks.map((t, idx) => unwrapTrackObject(t, idx))
          }, '*');
        } catch {}
      }
    }

    // VST3 Integration
    if (type === 'echo:powerpoint:toggleVstRack') {
      try {
        const btn = document.getElementById('echo-vst3-bar-btn');
        if (btn) {
          btn.click();
        } else {
          window.postMessage({ type: 'echo:vst3:toggleRack' }, '*');
        }
      } catch {}
    }

    if (type === 'echo:powerpoint:openVstManager') {
      try {
        window.postMessage({ type: 'echo:vst3:openPluginManager' }, '*');
      } catch {}
    }

    // Play Track with Debounce
    if (type === 'echo:powerpoint:playTrack' || type === 'echo:audition:playTrack') {
      const now = Date.now();
      if (now - lastPlayTrackTime < 180) return;
      lastPlayTrackTime = now;

      const { trackId, queueId } = e.data;
      const player = window.__echoExternalPlayer;
      const q = player?.queue ? (typeof player.queue === 'function' ? player.queue() : player.queue) : null;
      if (queueId && q?.playQueueItem) {
        try {
          await q.playQueueItem(queueId);
        } catch {}
        broadcastQueue(true);
        return;
      }
      if (trackId) {
        if (q?.items?.length) {
          const found = q.items.find(it => (it.track?.id === trackId || it.id === trackId));
          if (found && found.queueId && q.playQueueItem) {
            try {
              await q.playQueueItem(found.queueId);
            } catch {}
            broadcastQueue(true);
            return;
          }
        }
        if (q?.playTrack) {
          try {
            await q.playTrack(trackId);
          } catch {}
          broadcastQueue(true);
          return;
        }
        if (player?.playTrack) {
          try {
            await player.playTrack(trackId);
          } catch {}
          broadcastQueue(true);
          return;
        }
      }
    }

    // Command Dispatcher
    if (type === 'echo:powerpoint:command') {
      const { cmd, action, args, positionSeconds, time } = e.data;
      const targetCmd = cmd || action;
      const player = window.__echoExternalPlayer;
      try {
        if (targetCmd === 'playPause') {
          player?.playPause ? player.playPause() : window.echo?.playback?.playPause?.();
        } else if (targetCmd === 'next') {
          player?.next ? player.next() : window.echo?.playback?.next?.();
        } else if (targetCmd === 'previous') {
          player?.previous ? player.previous() : window.echo?.playback?.previous?.();
        } else if (targetCmd === 'seek') {
          const seekTime = Number(args?.time ?? args?.positionSeconds ?? positionSeconds ?? time ?? 0);
          if (window.echo?.playback?.seek) {
            await window.echo.playback.seek(seekTime);
          } else if (player?.seek) {
            await player.seek(seekTime);
          }
        } else if (targetCmd === 'setVolume') {
          const vol = Number(args?.volume ?? e.data.volume ?? 1);
          window.echo?.playback?.setVolume?.(vol);
        } else if (targetCmd === 'cycleRepeat') {
          window.echo?.playback?.cycleRepeat?.();
        } else if (targetCmd === 'toggleShuffle') {
          window.echo?.playback?.toggleShuffle?.();
        }
      } catch {}
    }

    // Cover Request from UI Sandbox
    if (type === 'echo:powerpoint:requestCover') {
      const { trackId, queueId, url, isThumb } = e.data;
      if (url) {
        resolveCoverDataUrl(url, isThumb !== false).then(thumbData => {
          if (thumbData) {
            broadcastQueueCover(trackId, queueId, url, thumbData);
          }
        }).catch(() => {});
      }
    }
  };

  window.addEventListener('message', onWindowMessage);

  // Host Status Listener with Deduplication
  let unsubAudioStatus = null;
  if (window.echo?.audio?.onStatus) {
    unsubAudioStatus = window.echo.audio.onStatus(async (status) => {
      broadcastAudioStatus(status);

      const trackChanged = status?.currentTrackId && status.currentTrackId !== lastBroadcastTrackId;
      const coverChanged = status?.currentTrackCoverUrl && status.currentTrackCoverUrl !== lastBroadcastCoverUrl;

      if (trackChanged || coverChanged) {
        lastBroadcastTrackId = status?.currentTrackId;
        lastBroadcastCoverUrl = status?.currentTrackCoverUrl;

        broadcastQueue(false);

        if (status?.currentTrackCoverUrl) {
          const albumUrl = status.currentTrackCoverUrl.replace('/thumb/', '/album/');
          const fullDataUrl = await resolveCoverDataUrl(albumUrl, false);
          if (fullDataUrl) {
            broadcastCover(status.currentTrackId, fullDataUrl);
          }
          const thumbDataUrl = await resolveCoverDataUrl(albumUrl, true);
          if (thumbDataUrl) {
            broadcastQueueCover(status.currentTrackId, null, albumUrl, thumbDataUrl);
          }
        }

        if (status?.currentTrackId) {
          const lyrics = await resolveLyrics(status.currentTrackId);
          if (lyrics) {
            broadcastLyrics(status.currentTrackId, lyrics);
          }
        }
      }
    });
  }

  // Real-time repository listeners
  let unsubQueueSession = null;
  let unsubQueueRepo = null;
  try {
    if (typeof window.echo?.playback?.onQueueSessionChanged === 'function') {
      unsubQueueSession = window.echo.playback.onQueueSessionChanged(() => {
        broadcastQueue(true);
      });
    }
    if (window.echo?.playback?.queueRepository?.onChanged) {
      unsubQueueRepo = window.echo.playback.queueRepository.onChanged(() => {
        broadcastQueue(true);
      });
    }
  } catch {}

  let unsubLyrics = null;
  try {
    if (window.echo?.lyrics?.onChanged) {
      unsubLyrics = window.echo.lyrics.onChanged(async (evt) => {
        const trackId = evt?.trackId || (await window.echo?.audio?.getStatus?.())?.currentTrackId;
        if (trackId) {
          const lyrics = await resolveLyrics(trackId);
          if (lyrics) broadcastLyrics(trackId, lyrics);
        }
      });
    }
  } catch {}

  // Periodic heartbeat (lightweight: audio status + queue check if empty)
  const heartbeatTimer = setInterval(() => {
    broadcastAudioStatus(null);
    if (!lastQueueSignature) {
      broadcastQueue(false);
    }
  }, 2000);

  window.__echoPowerPointBridgeDispose = () => {
    try { clearInterval(hostFixTimer); } catch {}
    try { clearInterval(heartbeatTimer); } catch {}
    try { window.removeEventListener('message', onWindowMessage); } catch {}
    try { if (typeof unsubAudioStatus === 'function') unsubAudioStatus(); } catch {}
    try { if (typeof unsubQueueSession === 'function') unsubQueueSession(); } catch {}
    try { if (typeof unsubQueueRepo === 'function') unsubQueueRepo(); } catch {}
    try { if (typeof unsubLyrics === 'function') unsubLyrics(); } catch {}
  };
})();
