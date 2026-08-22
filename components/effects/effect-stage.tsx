"use client";

import {
  ChangeEvent,
  CSSProperties,
  useCallback,
  useEffect,
  useMemo,
  useRef,
  useState,
} from "react";

type EffectType = "none" | "fireflies" | "leaves" | "film" | "roulette" | "gashapon" | "photos";

type ParticleSettings = {
  density: number;
  speed: number;
  size: number;
  opacity: number;
  glow: number;
  wind: number;
  color: string;
};

const defaultParticleSettings: ParticleSettings = {
  density: 90,
  speed: 0.8,
  size: 1,
  opacity: 0.85,
  glow: 0.8,
  wind: 0.15,
  color: "#ffd86b",
};

const effectLabels: Record<EffectType, string> = {
  none: "关闭特效",
  fireflies: "萤火虫",
  leaves: "枫叶飘落",
  film: "电影放映",
  roulette: "幸运转盘",
  gashapon: "扭蛋机",
  photos: "照片轮播",
};

type Particle = {
  x: number;
  y: number;
  vx: number;
  vy: number;
  size: number;
  alpha: number;
  phase: number;
  rotation: number;
  spin: number;
};

const random = (min: number, max: number) => min + Math.random() * (max - min);

function useCanvasParticles(
  canvasRef: React.RefObject<HTMLCanvasElement | null>,
  type: "fireflies" | "leaves",
  settings: ParticleSettings,
  active: boolean
) {
  const particlesRef = useRef<Particle[]>([]);
  const frameRef = useRef<number | null>(null);

  useEffect(() => {
    const canvas = canvasRef.current;
    if (!canvas || !active) return;
    const context = canvas.getContext("2d");
    if (!context) return;

    const resize = () => {
      const ratio = Math.min(window.devicePixelRatio || 1, 1.75);
      canvas.width = Math.max(1, Math.floor(window.innerWidth * ratio));
      canvas.height = Math.max(1, Math.floor(window.innerHeight * ratio));
      canvas.style.width = `${window.innerWidth}px`;
      canvas.style.height = `${window.innerHeight}px`;
      context.setTransform(ratio, 0, 0, ratio, 0, 0);
    };
    resize();
    window.addEventListener("resize", resize, { passive: true });

    const count = Math.max(8, Math.min(360, Math.round(settings.density)));
    particlesRef.current = Array.from({ length: count }, () => ({
      x: random(0, window.innerWidth),
      y: random(0, window.innerHeight),
      vx: random(-0.2, 0.2),
      vy: random(0.1, 0.7),
      size: random(2.5, 7) * settings.size,
      alpha: random(0.45, 1),
      phase: random(0, Math.PI * 2),
      rotation: random(0, Math.PI * 2),
      spin: random(-0.018, 0.018),
    }));

    let lastTime = performance.now();
    const render = (time: number) => {
      const dt = Math.min(2, (time - lastTime) / 16.67);
      lastTime = time;
      const width = window.innerWidth;
      const height = window.innerHeight;
      context.clearRect(0, 0, width, height);
      for (const particle of particlesRef.current) {
        particle.phase += 0.035 * settings.speed * dt;
        particle.rotation += particle.spin * settings.speed * dt;
        particle.x += (particle.vx + Math.sin(particle.phase) * settings.wind) * settings.speed * dt;
        particle.y += particle.vy * settings.speed * dt;
        if (particle.y > height + 30) {
          particle.y = -30;
          particle.x = random(0, width);
        }
        if (particle.x < -40) particle.x = width + 40;
        if (particle.x > width + 40) particle.x = -40;

        if (type === "fireflies") {
          const pulse = 0.55 + Math.sin(particle.phase * 1.7) * 0.35;
          const radius = Math.max(1, particle.size * pulse);
          context.save();
          context.globalAlpha = particle.alpha * settings.opacity * pulse;
          context.shadowColor = settings.color;
          context.shadowBlur = 15 * settings.glow * radius;
          context.fillStyle = settings.color;
          context.beginPath();
          context.arc(particle.x, particle.y, radius, 0, Math.PI * 2);
          context.fill();
          context.restore();
        } else {
          context.save();
          context.translate(particle.x, particle.y);
          context.rotate(particle.rotation);
          context.globalAlpha = particle.alpha * settings.opacity;
          context.fillStyle = settings.color;
          context.shadowColor = "rgba(77, 31, 4, .24)";
          context.shadowBlur = 3 * settings.glow;
          context.beginPath();
          context.moveTo(0, -particle.size * 1.5);
          context.bezierCurveTo(particle.size * 1.6, -particle.size, particle.size * 1.8, particle.size, 0, particle.size * 1.45);
          context.bezierCurveTo(-particle.size * 1.8, particle.size, -particle.size * 1.6, -particle.size, 0, -particle.size * 1.5);
          context.fill();
          context.strokeStyle = "rgba(89, 39, 10, .35)";
          context.lineWidth = Math.max(0.5, particle.size * 0.12);
          context.beginPath();
          context.moveTo(0, -particle.size * 1.15);
          context.lineTo(0, particle.size * 1.1);
          context.stroke();
          context.restore();
        }
      }
      frameRef.current = requestAnimationFrame(render);
    };
    frameRef.current = requestAnimationFrame(render);

    return () => {
      window.removeEventListener("resize", resize);
      if (frameRef.current !== null) cancelAnimationFrame(frameRef.current);
      context.clearRect(0, 0, window.innerWidth, window.innerHeight);
    };
  }, [active, canvasRef, settings, type]);
}

function CanvasEffect({ type, settings, active }: { type: "fireflies" | "leaves"; settings: ParticleSettings; active: boolean }) {
  const canvasRef = useRef<HTMLCanvasElement>(null);
  useCanvasParticles(canvasRef, type, settings, active);
  return <canvas ref={canvasRef} className="effect-stage__canvas" aria-hidden="true" />;
}

function FilmEffect({ active }: { active: boolean }) {
  if (!active) return null;
  return (
    <div className="effect-stage__film" aria-hidden="true">
      <div className="effect-stage__film-beam" />
      <div className="effect-stage__film-vignette" />
      <div className="effect-stage__film-scanlines" />
      <div className="effect-stage__film-dust" />
    </div>
  );
}

function RouletteEffect({ spinning, result }: { spinning: boolean; result: string }) {
  return (
    <div className={`effect-stage__roulette ${spinning ? "is-spinning" : ""}`} aria-live="polite">
      <div className="effect-stage__roulette-wheel">
        {Array.from({ length: 8 }, (_, index) => <span key={index}>{index + 1}</span>)}
      </div>
      <div className="effect-stage__roulette-pointer" />
      {!spinning && result && <div className="effect-stage__roulette-result">🎉 {result}</div>}
    </div>
  );
}

function GashaponEffect({ active, spinning }: { active: boolean; spinning: boolean }) {
  if (!active) return null;
  return (
    <div className="effect-stage__gashapon" aria-hidden="true">
      <div className={`effect-stage__capsule ${spinning ? "is-shaking" : ""}`}><i /><b /></div>
      <div className="effect-stage__machine"><span /><span /><span /><span /></div>
    </div>
  );
}

function PhotoEffect({ photos, index, active }: { photos: string[]; index: number; active: boolean }) {
  if (!active || photos.length === 0) return null;
  return (
    <div className="effect-stage__photo" aria-hidden="true">
      {/* eslint-disable-next-line @next/next/no-img-element */}
      <img src={photos[index % photos.length]} alt="" />
    </div>
  );
}

export default function EffectStage() {
  const [open, setOpen] = useState(false);
  const [present, setPresent] = useState(false);
  const [effect, setEffect] = useState<EffectType>("none");
  const [settings, setSettings] = useState<ParticleSettings>(defaultParticleSettings);
  const [playing, setPlaying] = useState(true);
  const [loop, setLoop] = useState(true);
  const [spinning, setSpinning] = useState(false);
  const [result, setResult] = useState("");
  const [photos, setPhotos] = useState<string[]>([]);
  const [photoIndex, setPhotoIndex] = useState(0);
  const [photoDuration, setPhotoDuration] = useState(4);
  const [quality, setQuality] = useState("high");
  const spinTimer = useRef<number | null>(null);

  const isParticleEffect = effect === "fireflies" || effect === "leaves";
  const isActive = playing && effect !== "none";
  const stageStyle = useMemo<CSSProperties>(() => ({ opacity: present ? 1 : 0.98 }), [present]);

  const selectEffect = (value: EffectType) => {
    setEffect(value);
    setResult("");
    setSpinning(false);
    if (value === "photos" && photos.length === 0) setOpen(true);
  };

  const startInteractive = () => {
    setResult("");
    setSpinning(true);
    if (spinTimer.current !== null) window.clearTimeout(spinTimer.current);
    spinTimer.current = window.setTimeout(() => {
      const prize = ["一等奖", "惊喜奖", "再来一次", "幸运奖", "特等奖"][Math.floor(Math.random() * 5)];
      setResult(prize);
      setSpinning(false);
    }, 3400);
  };

  const handlePhotos = (event: ChangeEvent<HTMLInputElement>) => {
    const files = Array.from(event.target.files || []).filter((file) => file.type.startsWith("image/"));
    setPhotos((previous) => {
      previous.forEach((url) => URL.revokeObjectURL(url));
      return files.map((file) => URL.createObjectURL(file));
    });
    setPhotoIndex(0);
    setEffect("photos");
    setPlaying(true);
  };

  useEffect(() => () => {
    photos.forEach((url) => URL.revokeObjectURL(url));
    if (spinTimer.current !== null) window.clearTimeout(spinTimer.current);
  }, [photos]);

  useEffect(() => {
    if (effect !== "photos" || photos.length < 2 || !playing) return;
    const timer = window.setInterval(() => {
      setPhotoIndex((current) => {
        const next = current + 1;
        if (next >= photos.length && !loop) return current;
        return next % photos.length;
      });
    }, Math.max(1, photoDuration) * 1000);
    return () => window.clearInterval(timer);
  }, [effect, loop, photoDuration, photos.length, playing]);

  const togglePresent = useCallback(async () => {
    const next = !present;
    setPresent(next);
    if (next && document.documentElement.requestFullscreen && !document.fullscreenElement) {
      try { await document.documentElement.requestFullscreen(); } catch { /* fullscreen may be blocked */ }
    } else if (!next && document.fullscreenElement && document.exitFullscreen) {
      try { await document.exitFullscreen(); } catch { /* already exited */ }
    }
  }, [present]);

  useEffect(() => {
    const onKeyDown = (event: KeyboardEvent) => {
      if (event.key === "F8") {
        event.preventDefault();
        void togglePresent();
      }
      if (event.key === "Escape" && present) setPresent(false);
    };
    window.addEventListener("keydown", onKeyDown);
    return () => window.removeEventListener("keydown", onKeyDown);
  }, [present, togglePresent]);

  return (
    <>
      <div className={`effect-stage ${present ? "is-presenting" : ""}`} style={stageStyle} aria-hidden={effect === "none"}>
        {isParticleEffect && <CanvasEffect type={effect} settings={{ ...settings, ...(quality === "low" ? { density: Math.min(settings.density, 45) } : {}) }} active={isActive} />}
        <FilmEffect active={effect === "film" && isActive} />
        {effect === "roulette" && isActive && <RouletteEffect spinning={spinning} result={result} />}
        <GashaponEffect active={effect === "gashapon" && isActive} spinning={spinning} />
        <PhotoEffect photos={photos} index={photoIndex} active={effect === "photos" && isActive} />
      </div>

      <button className={`effect-launcher ${open ? "is-open" : ""}`} type="button" onClick={() => setOpen((value) => !value)} title="打开网页特效">
        ✨<span>特效</span>
      </button>

      {open && <aside className="effect-panel" aria-label="网页特效控制面板">
        <div className="effect-panel__header">
          <div><strong>网页特效</strong><small>仅在本浏览器运行，不上传资源</small></div>
          <button type="button" onClick={() => setOpen(false)} aria-label="关闭">×</button>
        </div>
        <label className="effect-panel__field"><span>选择特效</span><select value={effect} onChange={(event) => selectEffect(event.target.value as EffectType)}>{Object.entries(effectLabels).map(([key, label]) => <option key={key} value={key}>{label}</option>)}</select></label>
        <div className="effect-panel__actions">
          <button type="button" onClick={() => setPlaying((value) => !value)}>{playing ? "暂停" : "播放"}</button>
          <button type="button" onClick={() => { setPlaying(false); requestAnimationFrame(() => setPlaying(true)); }}>重播</button>
          <button type="button" className="is-primary" onClick={() => void togglePresent()}>{present ? "退出演示" : "演示模式"}</button>
        </div>

        {isParticleEffect && <div className="effect-panel__section">
          <h3>粒子参数</h3>
          <label>数量 <input type="range" min="10" max="300" value={settings.density} onChange={(event) => setSettings((s) => ({ ...s, density: Number(event.target.value) }))} /><output>{settings.density}</output></label>
          <label>速度 <input type="range" min="0.1" max="3" step="0.1" value={settings.speed} onChange={(event) => setSettings((s) => ({ ...s, speed: Number(event.target.value) }))} /><output>{settings.speed.toFixed(1)}</output></label>
          <label>大小 <input type="range" min="0.4" max="2.5" step="0.1" value={settings.size} onChange={(event) => setSettings((s) => ({ ...s, size: Number(event.target.value) }))} /><output>{settings.size.toFixed(1)}</output></label>
          <label>风力 <input type="range" min="0" max="1" step="0.05" value={settings.wind} onChange={(event) => setSettings((s) => ({ ...s, wind: Number(event.target.value) }))} /><output>{settings.wind.toFixed(2)}</output></label>
          <label>透明度 <input type="range" min="0.1" max="1" step="0.05" value={settings.opacity} onChange={(event) => setSettings((s) => ({ ...s, opacity: Number(event.target.value) }))} /><output>{Math.round(settings.opacity * 100)}%</output></label>
          <label className="effect-panel__color">颜色 <input type="color" value={settings.color} onChange={(event) => setSettings((s) => ({ ...s, color: event.target.value }))} /></label>
        </div>}

        {effect === "photos" && <div className="effect-panel__section">
          <h3>照片播放</h3>
          <label className="effect-panel__upload">选择本地图片<input type="file" accept="image/*" multiple onChange={handlePhotos} /></label>
          <label>间隔（秒） <input type="range" min="1" max="20" value={photoDuration} onChange={(event) => setPhotoDuration(Number(event.target.value))} /><output>{photoDuration}</output></label>
          <label className="effect-panel__check"><input type="checkbox" checked={loop} onChange={(event) => setLoop(event.target.checked)} /> 循环播放</label>
          <small className="effect-panel__hint">{photos.length ? `已选择 ${photos.length} 张图片` : "图片只保存在当前浏览器"}</small>
        </div>}

        {(effect === "roulette" || effect === "gashapon") && <div className="effect-panel__section">
          <h3>互动控制</h3>
          <button type="button" className="effect-panel__big-action" onClick={startInteractive}>{effect === "roulette" ? (spinning ? "抽取中…" : "开始抽奖") : (spinning ? "扭蛋中…" : "打开扭蛋")}</button>
          {result && <p className="effect-panel__result">结果：{result}</p>}
        </div>}

        <label className="effect-panel__field"><span>渲染质量</span><select value={quality} onChange={(event) => setQuality(event.target.value)}><option value="high">高</option><option value="low">省电/移动端</option></select></label>
        <p className="effect-panel__footer">快捷键：F8 进入/退出演示模式 · Esc 退出</p>
      </aside>}
    </>
  );
}
