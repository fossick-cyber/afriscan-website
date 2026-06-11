// AfriScan hero — real-time WebGL diorama (Three.js).
// Falls back silently to the CSS scene when WebGL/CDN/reduced-motion rule it out.

const sceneRoot = document.querySelector('.hero-scene');
const reduced = window.matchMedia('(prefers-reduced-motion: reduce)').matches;

function webglOK() {
  try {
    const c = document.createElement('canvas');
    return !!(c.getContext('webgl2') || c.getContext('webgl'));
  } catch { return false; }
}

if (sceneRoot && !reduced && webglOK()) boot();

async function boot() {
  let THREE, RoomEnvironment;
  try {
    THREE = await import('three');
    ({ RoomEnvironment } = await import('three/addons/environments/RoomEnvironment.js'));
  } catch (e) {
    console.warn('3D scene unavailable, keeping CSS fallback:', e);
    return;
  }

  // optional post-processing — failure here just disables bloom, scene still renders
  let PP = null;
  try {
    const mods = await Promise.all([
      import('three/addons/postprocessing/EffectComposer.js'),
      import('three/addons/postprocessing/RenderPass.js'),
      import('three/addons/postprocessing/UnrealBloomPass.js'),
      import('three/addons/postprocessing/OutputPass.js'),
      import('three/addons/postprocessing/SMAAPass.js'),
    ]);
    PP = {
      EffectComposer: mods[0].EffectComposer, RenderPass: mods[1].RenderPass,
      UnrealBloomPass: mods[2].UnrealBloomPass, OutputPass: mods[3].OutputPass,
      SMAAPass: mods[4].SMAAPass,
    };
  } catch (e) { console.warn('post-processing unavailable:', e); }

  // ---------- renderer / scene / camera ----------
  const canvas = document.createElement('canvas');
  canvas.className = 'scene3d-canvas';
  sceneRoot.appendChild(canvas);

  const renderer = new THREE.WebGLRenderer({ canvas, antialias: true, alpha: true });
  const isMobile = window.matchMedia('(max-width: 700px)').matches;
  renderer.setPixelRatio(Math.min(window.devicePixelRatio, isMobile ? 1.5 : 2));
  renderer.shadowMap.enabled = !isMobile;
  renderer.shadowMap.type = THREE.PCFSoftShadowMap;
  renderer.toneMapping = THREE.ACESFilmicToneMapping;
  renderer.toneMappingExposure = 1.0;

  const scene = new THREE.Scene();
  scene.fog = new THREE.FogExp2(0x0d1b29, 0.0105);

  // gradient sky dome (dark zenith → faint teal/orange horizon)
  {
    const skyGeo = new THREE.SphereGeometry(160, 32, 16);
    const skyMat = new THREE.ShaderMaterial({
      side: THREE.BackSide, depthWrite: false, fog: false,
      uniforms: {
        top: { value: new THREE.Color(0x070d16) },
        mid: { value: new THREE.Color(0x123048) },
        hor: { value: new THREE.Color(0x1d4257) },
      },
      vertexShader: 'varying vec3 vP; void main(){ vP = position; gl_Position = projectionMatrix * modelViewMatrix * vec4(position,1.0); }',
      fragmentShader: `varying vec3 vP; uniform vec3 top; uniform vec3 mid; uniform vec3 hor;
        void main(){ float h = normalize(vP).y;
          vec3 c = mix(hor, mid, smoothstep(-0.02, 0.32, h));
          c = mix(c, top, smoothstep(0.18, 0.7, h));
          gl_FragColor = vec4(c, 1.0); }`,
    });
    scene.add(new THREE.Mesh(skyGeo, skyMat));
  }

  const camera = new THREE.PerspectiveCamera(44, 2, 0.1, 320);
  camera.position.set(0, 12.5, 23.5);

  const pmrem = new THREE.PMREMGenerator(renderer);
  scene.environment = pmrem.fromScene(new RoomEnvironment(), 0.04).texture;

  // ---------- lights ----------
  scene.add(new THREE.HemisphereLight(0x24405c, 0x0b1016, 0.55));
  const moon = new THREE.DirectionalLight(0xbfd4e6, 1.15);
  moon.position.set(-22, 28, 16);
  moon.castShadow = !isMobile;
  moon.shadow.mapSize.set(2048, 2048);
  Object.assign(moon.shadow.camera, { left: -45, right: 45, top: 40, bottom: -25, far: 90 });
  moon.shadow.bias = -0.0004;
  scene.add(moon);
  const warm = new THREE.DirectionalLight(0xe8672f, 0.18);
  warm.position.set(24, 6, -14);
  scene.add(warm);

  // ---------- helpers ----------
  const texCanvas = (w, h, draw) => {
    const c = document.createElement('canvas'); c.width = w; c.height = h;
    draw(c.getContext('2d'));
    const t = new THREE.CanvasTexture(c); t.anisotropy = 4; return t;
  };
  const radialSprite = (color, inner = 1) => texCanvas(128, 128, ctx => {
    const g = ctx.createRadialGradient(64, 64, 2, 64, 64, 64);
    g.addColorStop(0, color.replace('A', String(inner)));
    g.addColorStop(1, color.replace('A', '0'));
    ctx.fillStyle = g; ctx.fillRect(0, 0, 128, 128);
  });

  // ---------- terrain ----------
  const groundGeo = new THREE.PlaneGeometry(170, 90, 120, 60);
  {
    const pos = groundGeo.attributes.position;
    const colors = [];
    const base = new THREE.Color(0x122b3d), grass = new THREE.Color(0x274a33), dirt = new THREE.Color(0x4a3b26);
    for (let i = 0; i < pos.count; i++) {
      const x = pos.getX(i), y = pos.getY(i);           // plane local: y → world -z
      const nearPipe = Math.abs(y) < 3;
      const h = nearPipe ? 0 :
        (Math.sin(x * 0.18) * Math.cos(y * 0.22) + Math.sin(x * 0.07 + y * 0.13) * 1.6) * 0.22;
      pos.setZ(i, h);
      const m = (Math.sin(x * 0.35 + 9) * Math.cos(y * 0.3 + 2) + 1) / 2;
      const c = base.clone().lerp(m > 0.62 ? grass : dirt, Math.max(0, m - 0.45) * 0.55);
      colors.push(c.r, c.g, c.b);
    }
    groundGeo.setAttribute('color', new THREE.Float32BufferAttribute(colors, 3));
    groundGeo.computeVertexNormals();
  }
  const ground = new THREE.Mesh(groundGeo,
    new THREE.MeshStandardMaterial({ vertexColors: true, roughness: 0.95, metalness: 0 }));
  ground.rotation.x = -Math.PI / 2;
  ground.position.z = 4;
  ground.receiveShadow = true;
  scene.add(ground);

  const grid = new THREE.GridHelper(170, 68, 0x5eead4, 0x5eead4);
  grid.material.transparent = true; grid.material.opacity = 0.07;
  grid.position.y = 0.03; grid.position.z = 4;
  scene.add(grid);

  // ---------- buffer zones (flat translucent bands beside the pipe) ----------
  const zone = (depth0, depth1, color, op) => {
    const m = new THREE.Mesh(new THREE.PlaneGeometry(170, depth1 - depth0),
      new THREE.MeshBasicMaterial({ color, transparent: true, opacity: op, depthWrite: false }));
    m.rotation.x = -Math.PI / 2;
    m.position.set(0, 0.05, (depth0 + depth1) / 2);
    scene.add(m);
  };
  zone(0.55, 1.67, 0xef4444, 0.10);  // 0–50 m
  zone(1.67, 5.0, 0xf5c518, 0.05);   // 50–150 m

  // ---------- gas pipeline along X ----------
  // brushed-metal normal map (fine lengthwise streaks) + roughness variation
  const steelNormal = texCanvas(512, 64, ctx => {
    ctx.fillStyle = '#8080ff'; ctx.fillRect(0, 0, 512, 64);
    for (let i = 0; i < 900; i++) {
      const y = Math.random() * 64, n = 110 + Math.random() * 90;
      ctx.strokeStyle = `rgba(${n},${n},255,0.5)`; ctx.lineWidth = Math.random() * 1.2;
      ctx.beginPath(); ctx.moveTo(Math.random() * 512, y); ctx.lineTo(Math.random() * 512, y); ctx.stroke();
    }
  });
  steelNormal.wrapS = steelNormal.wrapT = THREE.RepeatWrapping; steelNormal.repeat.set(40, 2);
  const steelRough = texCanvas(256, 64, ctx => {
    ctx.fillStyle = '#4a4a4a'; ctx.fillRect(0, 0, 256, 64);
    for (let i = 0; i < 200; i++) { ctx.fillStyle = `rgba(${Math.random()>.5?200:30},0,0,0.18)`;
      ctx.fillRect(Math.random()*256, Math.random()*64, 6+Math.random()*30, 2+Math.random()*6); }
  });
  steelRough.wrapS = THREE.RepeatWrapping; steelRough.repeat.set(20, 1);
  const steel = new THREE.MeshStandardMaterial({
    color: 0x8c98a6, metalness: 0.35, roughness: 0.82,
    normalMap: steelNormal, normalScale: new THREE.Vector2(0.25, 0.25),
    roughnessMap: steelRough, envMapIntensity: 0.35,
  });
  const pipe = new THREE.Mesh(new THREE.CylinderGeometry(0.45, 0.45, 170, 48), steel);
  pipe.rotation.z = Math.PI / 2;
  pipe.position.y = 0.78;
  pipe.castShadow = true; pipe.receiveShadow = true;
  scene.add(pipe);
  const weldMat = new THREE.MeshStandardMaterial({ color: 0x55606c, metalness: 0.85, roughness: 0.5 });
  const sleeperMat = new THREE.MeshStandardMaterial({ color: 0x8d8d86, roughness: 0.9 });
  for (let x = -80; x <= 80; x += 6) {
    const weld = new THREE.Mesh(new THREE.TorusGeometry(0.46, 0.025, 8, 24), weldMat);
    weld.rotation.y = Math.PI / 2; weld.position.set(x + 3, 0.78, 0);
    scene.add(weld);
    const sleeper = new THREE.Mesh(new THREE.BoxGeometry(0.5, 0.36, 1.1), sleeperMat);
    sleeper.position.set(x, 0.18, 0); sleeper.castShadow = true;
    scene.add(sleeper);
  }
  // service road on the far side
  const road = new THREE.Mesh(new THREE.PlaneGeometry(170, 1.4),
    new THREE.MeshStandardMaterial({ color: 0x6b5836, roughness: 1 }));
  road.rotation.x = -Math.PI / 2; road.position.set(0, 0.04, -1.9);
  road.receiveShadow = true;
  scene.add(road);

  // gas marker sign
  const signTex = texCanvas(256, 128, ctx => {
    ctx.fillStyle = '#f4c430'; ctx.fillRect(0, 0, 256, 128);
    ctx.strokeStyle = '#1a1303'; ctx.lineWidth = 10; ctx.strokeRect(5, 5, 246, 118);
    ctx.fillStyle = '#1a1303'; ctx.textAlign = 'center'; ctx.font = '900 34px Arial';
    ctx.fillText('⚠ GAS', 128, 52); ctx.fillText('PIPELINE', 128, 96);
  });
  const post = new THREE.Mesh(new THREE.CylinderGeometry(0.04, 0.04, 1.5, 8),
    new THREE.MeshStandardMaterial({ color: 0xb9b9b9, metalness: 0.6, roughness: 0.4 }));
  post.position.set(-8.5, 0.75, 1.4); scene.add(post);
  const sign = new THREE.Mesh(new THREE.PlaneGeometry(1.3, 0.65),
    new THREE.MeshStandardMaterial({ map: signTex, roughness: 0.6 }));
  sign.position.set(-8.5, 1.65, 1.42); scene.add(sign);

  // ---------- huts / trees ----------
  const thatchTex = texCanvas(128, 128, ctx => {
    ctx.fillStyle = '#8a6c38'; ctx.fillRect(0, 0, 128, 128);
    for (let i = 0; i < 240; i++) {
      ctx.strokeStyle = `rgba(60,45,18,${0.15 + Math.random() * 0.3})`;
      const x = Math.random() * 128; ctx.beginPath();
      ctx.moveTo(x, Math.random() * 20); ctx.lineTo(x + (Math.random() * 6 - 3), 128); ctx.stroke();
    }
  });
  thatchTex.wrapS = thatchTex.wrapT = THREE.RepeatWrapping; thatchTex.repeat.set(4, 1);
  const wallTex = texCanvas(128, 64, ctx => {
    ctx.fillStyle = '#9c6b42'; ctx.fillRect(0, 0, 128, 64);
    for (let i = 0; i < 26; i++) {
      ctx.fillStyle = `rgba(216,167,107,${Math.random() * 0.16})`;
      ctx.beginPath(); ctx.ellipse(Math.random() * 128, Math.random() * 64, 8 + Math.random() * 10, 4 + Math.random() * 5, 0, 0, 7); ctx.fill();
    }
  });
  wallTex.wrapS = THREE.RepeatWrapping; wallTex.repeat.set(3, 1);

  function hut(scale = 1) {
    const g = new THREE.Group();
    const wall = new THREE.Mesh(new THREE.CylinderGeometry(0.55, 0.6, 0.55, 14),
      new THREE.MeshStandardMaterial({ map: wallTex, roughness: 0.95 }));
    wall.position.y = 0.275;
    const roof = new THREE.Mesh(new THREE.ConeGeometry(0.85, 0.62, 14),
      new THREE.MeshStandardMaterial({ map: thatchTex, roughness: 1 }));
    roof.position.y = 0.86;
    const door = new THREE.Mesh(new THREE.BoxGeometry(0.22, 0.34, 0.05),
      new THREE.MeshStandardMaterial({ color: 0x2a1c0e, roughness: 1 }));
    door.position.set(0, 0.18, 0.58);
    g.add(wall, roof, door);
    g.traverse(o => { if (o.isMesh) { o.castShadow = true; o.receiveShadow = true; } });
    g.scale.setScalar(scale);
    return g;
  }
  function tree(scale = 1) {
    const g = new THREE.Group();
    const trunk = new THREE.Mesh(new THREE.CylinderGeometry(0.06, 0.1, 0.9, 6),
      new THREE.MeshStandardMaterial({ color: 0x4a3520, roughness: 1 }));
    trunk.position.y = 0.45;
    const canopyMat = new THREE.MeshStandardMaterial({ color: 0x2c4a28, roughness: 1 });
    const c1 = new THREE.Mesh(new THREE.SphereGeometry(0.62, 10, 6), canopyMat);
    c1.scale.set(1.5, 0.45, 1.2); c1.position.y = 1.05;
    const c2 = new THREE.Mesh(new THREE.SphereGeometry(0.4, 8, 5),
      new THREE.MeshStandardMaterial({ color: 0x3a6132, roughness: 1 }));
    c2.scale.set(1.3, 0.4, 1); c2.position.set(0.3, 1.2, 0.1);
    g.add(trunk, c1, c2);
    g.traverse(o => { if (o.isMesh) o.castShadow = true; });
    g.scale.setScalar(scale);
    return g;
  }

  // settlements at true distances (1 unit = 30 m); zone by z: red <1.67, amber <5, else green
  const settlements = [
    { id: 1, x: -15, z: 8.6, dist: '258m', cls: 'green', huts: [[0, 0, 1], [1.6, .5, .8]], treeAt: [-1.5, .4] },
    { id: 2, x: -8,  z: 5.6, dist: '168m', cls: 'green', huts: [[0, 0, .95]], treeAt: [-1.4, .3] },
    { id: 3, x: -2,  z: 3.3, dist: '99m',  cls: 'amber', huts: [[0, 0, 1], [1.4, .4, .78]], treeAt: null, mobOff: 26 },
    { id: 4, x: 5,   z: 1.05, dist: '32m', cls: 'red',   huts: [[0, 0, 1], [1.3, .5, .82]], treeAt: null },
    { id: 5, x: 11,  z: 2.5, dist: '75m',  cls: 'amber', huts: [[0, 0, .9]], treeAt: [1.5, .3], mobOff: -22 },
    { id: 6, x: 17,  z: 6.4, dist: '192m', cls: 'green', huts: [[0, 0, 1], [-1.4, .4, .8]], treeAt: [1.6, .5] },
  ];
  const detColors = { green: 0x22c55e, amber: 0xf5c518, red: 0xef4444 };

  for (const s of settlements) {
    s.group = new THREE.Group();
    s.group.position.set(s.x, 0, s.z);
    s.hutGroup = new THREE.Group();
    for (const [hx, hz, hs] of s.huts) {
      const h = hut(hs); h.position.set(hx, 0, hz); s.hutGroup.add(h);
    }
    s.group.add(s.hutGroup);
    if (s.treeAt) { const t = tree(0.9); t.position.set(s.treeAt[0], 0, s.treeAt[1]); s.group.add(t); }
    scene.add(s.group);
    // detection box from TRUE world bounds (matrices updated first)
    s.group.updateMatrixWorld(true);
    const bbox = new THREE.Box3().setFromObject(s.hutGroup);
    const size = bbox.getSize(new THREE.Vector3()).addScalar(0.5);
    const center = bbox.getCenter(new THREE.Vector3());
    s.det = new THREE.LineSegments(
      new THREE.EdgesGeometry(new THREE.BoxGeometry(size.x, size.y, size.z)),
      new THREE.LineBasicMaterial({ color: detColors[s.cls], transparent: true, opacity: 0 }));
    s.det.position.copy(center);                       // world space (child of scene)
    scene.add(s.det);
    s.labelPos = new THREE.Vector3(center.x, center.y + size.y / 2 + 0.55, center.z);
    // DOM label
    s.tag = document.createElement('div');
    s.tag.className = `tag3d tag3d-${s.cls}`;
    s.tag.innerHTML = (s.cls === 'red' ? '<i class="warn">!</i>' : `<i class="led${s.cls === 'amber' ? ' led-pulse' : ''}"></i>`) +
      `BLDG-00${s.id} &middot; ${s.dist}`;
    s.tag.style.opacity = '0';
    sceneRoot.appendChild(s.tag);
  }
  const bldg3 = settlements.find(s => s.cls === 'red');   // demolition target

  // scattered background trees
  for (const [tx, tz, ts] of [[-22, 4.5, 0.9], [21, 5, 0.85], [-19, 9.5, 1.0], [14, 9.5, 0.8], [24, 2.5, 0.8], [-24, 2, 0.75], [2, 11.5, 0.7], [9, 12, 0.75]]) {
    const t = tree(ts); t.position.set(tx, 0, tz); scene.add(t);
  }

  // ground dressing: rocks + shrubs (instanced), seeded pseudo-random
  let seed = 1337; const rnd = () => { seed = (seed * 16807) % 2147483647; return seed / 2147483647; };
  {
    const rockGeo = new THREE.IcosahedronGeometry(0.22, 0);
    const rockMat = new THREE.MeshStandardMaterial({ color: 0x6b6b66, roughness: 1, flatShading: true });
    const rocks = new THREE.InstancedMesh(rockGeo, rockMat, 26);
    const shrubGeo = new THREE.IcosahedronGeometry(0.3, 0);
    const shrubMat = new THREE.MeshStandardMaterial({ color: 0x35502c, roughness: 1, flatShading: true });
    const shrubs = new THREE.InstancedMesh(shrubGeo, shrubMat, 30);
    const m = new THREE.Matrix4(), q = new THREE.Quaternion(), s = new THREE.Vector3(), pos = new THREE.Vector3();
    for (let i = 0; i < 26; i++) {
      pos.set((rnd() - 0.5) * 46, 0.12, 1 + rnd() * 13);
      q.setFromEuler(new THREE.Euler(rnd(), rnd() * 6, rnd()));
      const sc = 0.4 + rnd() * 1.1; s.set(sc, sc * (0.6 + rnd() * 0.5), sc);
      rocks.setMatrixAt(i, m.compose(pos, q, s));
    }
    for (let i = 0; i < 30; i++) {
      pos.set((rnd() - 0.5) * 50, 0.16, 0.8 + rnd() * 14);
      q.setFromEuler(new THREE.Euler(0, rnd() * 6, 0));
      const sc = 0.5 + rnd() * 0.9; s.set(sc * 1.4, sc * 0.7, sc * 1.4);
      shrubs.setMatrixAt(i, m.compose(pos, q, s));
    }
    rocks.castShadow = shrubs.castShadow = true;
    rocks.receiveShadow = shrubs.receiveShadow = true;
    scene.add(rocks, shrubs);
  }

  // cattle kraal — a ring of stick fence posts beside the far settlement
  {
    const kraal = new THREE.Group();
    const postMat = new THREE.MeshStandardMaterial({ color: 0x5a4327, roughness: 1 });
    const R = 1.5;
    for (let a = 0; a < Math.PI * 2; a += Math.PI / 11) {
      const p = new THREE.Mesh(new THREE.CylinderGeometry(0.035, 0.045, 0.6, 5), postMat);
      p.position.set(Math.cos(a) * R, 0.3, Math.sin(a) * R);
      p.rotation.z = (rnd() - 0.5) * 0.2; p.castShadow = true;
      kraal.add(p);
    }
    const rail = new THREE.Mesh(new THREE.TorusGeometry(R, 0.03, 5, 28), postMat);
    rail.rotation.x = Math.PI / 2; rail.position.y = 0.45; kraal.add(rail);
    kraal.position.set(-13.5, 0, 11.2);
    scene.add(kraal);
  }

  // dirt footpaths linking the settlements to the corridor
  const pathMat = new THREE.MeshStandardMaterial({ color: 0x6b5836, roughness: 1, transparent: true, opacity: 0.6 });
  for (const s of settlements) {
    const len = Math.hypot(s.x - s.x, s.z - 0.6) + s.z;
    const path = new THREE.Mesh(new THREE.PlaneGeometry(0.5, s.z + 0.6), pathMat);
    path.rotation.x = -Math.PI / 2; path.position.set(s.x, 0.045, (s.z + 0.6) / 2 + 0.3);
    scene.add(path);
  }

  // ---------- monitoring station (radar) ----------
  const mast = new THREE.Mesh(new THREE.CylinderGeometry(0.05, 0.07, 1.8, 6),
    new THREE.MeshStandardMaterial({ color: 0x8a98ab, metalness: 0.7, roughness: 0.4 }));
  mast.position.set(-20, 0.9, 1.2); scene.add(mast);
  const rings = [];
  for (let i = 0; i < 2; i++) {
    const r = new THREE.Mesh(new THREE.RingGeometry(0.96, 1, 40),
      new THREE.MeshBasicMaterial({ color: 0x5eead4, transparent: true, opacity: 0.5, side: THREE.DoubleSide, depthWrite: false }));
    r.rotation.x = -Math.PI / 2; r.position.set(-20, 0.06, 1.2);
    scene.add(r); rings.push(r);
  }

  // ---------- drone ----------
  function buildDrone() {
    const g = new THREE.Group();
    const bodyMat = new THREE.MeshStandardMaterial({ color: 0x39485d, metalness: 0.45, roughness: 0.4 });
    const darkMat = new THREE.MeshStandardMaterial({ color: 0x1d2735, metalness: 0.4, roughness: 0.5 });
    const body = new THREE.Mesh(new THREE.SphereGeometry(0.55, 18, 12), bodyMat);
    body.scale.set(1.5, 0.42, 0.8);
    const canopy = new THREE.Mesh(new THREE.SphereGeometry(0.34, 14, 10), new THREE.MeshStandardMaterial({ color: 0x55667e, metalness: 0.6, roughness: 0.25 }));
    canopy.scale.set(1.1, 0.5, 0.75); canopy.position.y = 0.16;
    const stripe = new THREE.Mesh(new THREE.TorusGeometry(0.56, 0.035, 6, 24), new THREE.MeshStandardMaterial({ color: 0xe8672f, roughness: 0.4 }));
    stripe.rotation.x = Math.PI / 2; stripe.scale.set(1.45, 0.78, 1);
    g.add(body, canopy, stripe);
    g.rotors = [];
    for (const [ax, az] of [[-0.95, -0.62], [0.95, -0.62], [-0.95, 0.62], [0.95, 0.62]]) {
      const arm = new THREE.Mesh(new THREE.CylinderGeometry(0.045, 0.045, Math.hypot(ax, az), 6), darkMat);
      arm.rotation.z = Math.PI / 2;
      arm.rotation.y = -Math.atan2(az, ax);
      arm.position.set(ax / 2, 0.02, az / 2);
      const pod = new THREE.Mesh(new THREE.CylinderGeometry(0.09, 0.11, 0.16, 8), darkMat);
      pod.position.set(ax, 0.08, az);
      const rotor = new THREE.Mesh(new THREE.CylinderGeometry(0.5, 0.5, 0.012, 18),
        new THREE.MeshBasicMaterial({ color: 0xc8ebe4, transparent: true, opacity: 0.16, depthWrite: false }));
      rotor.position.set(ax, 0.18, az);
      g.add(arm, pod, rotor);
      g.rotors.push(rotor);
    }
    // gimbal camera
    const gim = new THREE.Mesh(new THREE.SphereGeometry(0.14, 10, 8), darkMat);
    gim.position.set(0.42, -0.22, 0);
    const lens = new THREE.Mesh(new THREE.SphereGeometry(0.07, 10, 8),
      new THREE.MeshStandardMaterial({ color: 0x9bfff0, emissive: 0x5eead4, emissiveIntensity: 2.6, roughness: 0.05 }));
    lens.position.set(0.52, -0.24, 0);
    g.add(gim, lens);
    // nav lights (emissive + glow sprites)
    const glowTex = radialSprite('rgba(255,80,80,A)');
    const glowTexG = radialSprite('rgba(80,255,140,A)');
    g.navL = new THREE.Sprite(new THREE.SpriteMaterial({ map: glowTex, transparent: true, depthWrite: false }));
    g.navL.scale.setScalar(0.55); g.navL.position.set(-0.8, 0, -0.62);
    g.navR = new THREE.Sprite(new THREE.SpriteMaterial({ map: glowTexG, transparent: true, depthWrite: false }));
    g.navR.scale.setScalar(0.55); g.navR.position.set(-0.8, 0, 0.62);
    g.add(g.navL, g.navR);
    g.traverse(o => { if (o.isMesh) o.castShadow = true; });
    return g;
  }
  const drone = buildDrone();
  scene.add(drone);

  // scan cone + ground spot
  const coneH = 4.9;
  const cone = new THREE.Mesh(new THREE.ConeGeometry(2.1, coneH, 32, 1, true),
    new THREE.MeshBasicMaterial({ color: 0x5eead4, transparent: true, opacity: 0.12, depthWrite: false, side: THREE.DoubleSide, blending: THREE.AdditiveBlending }));
  scene.add(cone);
  const spot = new THREE.Mesh(new THREE.CircleGeometry(2.1, 32),
    new THREE.MeshBasicMaterial({ color: 0x5eead4, transparent: true, opacity: 0.16, depthWrite: false, blending: THREE.AdditiveBlending }));
  spot.rotation.x = -Math.PI / 2; spot.position.y = 0.07;
  scene.add(spot);
  // bright scan ring on the ground for the bloom to catch
  const scanRing = new THREE.Mesh(new THREE.RingGeometry(1.85, 2.1, 40),
    new THREE.MeshBasicMaterial({ color: 0x7ef7e6, transparent: true, opacity: 0.6, depthWrite: false, blending: THREE.AdditiveBlending, side: THREE.DoubleSide }));
  scanRing.rotation.x = -Math.PI / 2; scanRing.position.y = 0.08;
  scene.add(scanRing);

  // ---------- satellite ----------
  const sat = new THREE.Group();
  {
    const bus = new THREE.Mesh(new THREE.BoxGeometry(0.8, 0.8, 0.8),
      new THREE.MeshStandardMaterial({ color: 0xa87b22, metalness: 0.7, roughness: 0.35 }));
    const wingMat = new THREE.MeshStandardMaterial({ color: 0x1f63b8, metalness: 0.6, roughness: 0.3 });
    const w1 = new THREE.Mesh(new THREE.BoxGeometry(2.4, 0.05, 0.9), wingMat); w1.position.x = -1.7;
    const w2 = w1.clone(); w2.position.x = 1.7;
    sat.add(bus, w1, w2);
  }
  sat.scale.setScalar(0.8);
  scene.add(sat);

  // ---------- stars ----------
  {
    const n = 350, pts = [];
    for (let i = 0; i < n; i++) {
      const a = Math.random() * Math.PI, r = 120 + Math.random() * 40;
      pts.push(Math.cos(a) * r * (Math.random() > .5 ? 1 : -1), 20 + Math.random() * 70, -40 - Math.random() * 80);
    }
    const geo = new THREE.BufferGeometry();
    geo.setAttribute('position', new THREE.Float32BufferAttribute(pts, 3));
    scene.add(new THREE.Points(geo, new THREE.PointsMaterial({ color: 0xbfd9e8, size: 0.35, transparent: true, opacity: 0.8 })));
  }

  // ---------- bulldozer ----------
  function buildDozer() {
    const g = new THREE.Group();
    const yellow = new THREE.MeshStandardMaterial({ color: 0xc98a1e, metalness: 0.3, roughness: 0.55 });
    const dark = new THREE.MeshStandardMaterial({ color: 0x2b2f36, roughness: 0.8 });
    const body = new THREE.Mesh(new THREE.BoxGeometry(1.7, 0.55, 1.05), yellow); body.position.y = 0.75;
    const cab = new THREE.Mesh(new THREE.BoxGeometry(0.7, 0.55, 0.85), yellow); cab.position.set(-0.25, 1.3, 0);
    const glass = new THREE.Mesh(new THREE.BoxGeometry(0.6, 0.4, 0.75),
      new THREE.MeshStandardMaterial({ color: 0x9fd4e8, metalness: 0.2, roughness: 0.1 }));
    glass.position.set(-0.25, 1.32, 0);
    const trackL = new THREE.Mesh(new THREE.BoxGeometry(1.9, 0.5, 0.34), dark); trackL.position.set(0, 0.3, 0.56);
    const trackR = trackL.clone(); trackR.position.z = -0.56;
    const blade = new THREE.Mesh(new THREE.CylinderGeometry(0.55, 0.55, 1.5, 12, 1, true, 0, Math.PI * 0.8),
      new THREE.MeshStandardMaterial({ color: 0x8a98ab, metalness: 0.8, roughness: 0.35, side: THREE.DoubleSide }));
    blade.rotation.z = Math.PI / 2; blade.rotation.y = Math.PI / 2;
    blade.position.set(-1.35, 0.55, 0);
    const pipeEx = new THREE.Mesh(new THREE.CylinderGeometry(0.05, 0.05, 0.5, 6), dark);
    pipeEx.position.set(0.3, 1.35, 0.3);
    g.add(body, cab, glass, trackL, trackR, blade, pipeEx);
    g.traverse(o => { if (o.isMesh) o.castShadow = true; });
    return g;
  }
  const dozer = buildDozer();
  dozer.position.set(46, 0, bldg3.z);
  dozer.rotation.y = Math.PI;        // blade faces -x (toward the village)
  dozer.visible = false;
  scene.add(dozer);

  // dust sprites for demolition
  const dustTex = radialSprite('rgba(196,170,130,A)', 0.55);
  const dusts = [];
  for (let i = 0; i < 7; i++) {
    const d = new THREE.Sprite(new THREE.SpriteMaterial({ map: dustTex, transparent: true, opacity: 0, depthWrite: false }));
    d.scale.setScalar(1.4);
    scene.add(d); dusts.push(d);
  }

  // ---------- swap CSS scene for WebGL ----------
  sceneRoot.classList.add('webgl-on');

  // ---------- animation ----------
  const clock = new THREE.Clock();
  const LOOP = 18, X0 = -38, X1 = 38;
  let parX = 0, parY = 0, tgtParX = 0, tgtParY = 0;
  const hero = document.querySelector('.hero');
  if (hero && window.matchMedia('(pointer: fine)').matches) {
    hero.addEventListener('mousemove', e => {
      const r = hero.getBoundingClientRect();
      tgtParX = ((e.clientX - r.left) / r.width - 0.5) * 3.2;
      tgtParY = ((e.clientY - r.top) / r.height - 0.5) * 1.6;
    });
    hero.addEventListener('mouseleave', () => { tgtParX = 0; tgtParY = 0; });
  }

  // demolition state
  let demo = null;           // {t0}
  let running = false;
  const DEMO = { driveIn: 3.2, push: 1.1, hold: 0.8, out: 2.6, regrow: 1.4 };
  function demolish() {
    if (running) return;
    running = true;
    demo = { t0: clock.getElapsedTime() };
    dozer.visible = true;
  }
  window.__afriscanDemolish3D = demolish;
  window.__afriscanState = () => ({ running, taps });
  window.__afriscanBldg3Rect = () => bldg3.tag.getBoundingClientRect();

  // secret triggers (keyword + triple-tap on red label/buildings)
  let buf = '';
  document.addEventListener('keydown', e => {
    if (e.key.length !== 1 || e.metaKey || e.ctrlKey || e.altKey) return;
    const tag = (e.target.tagName || '').toLowerCase();
    if (tag === 'input' || tag === 'textarea' || tag === 'select') return;
    buf = (buf + e.key.toLowerCase()).slice(-4);
    if (buf === 'doze') demolish();
  });
  let taps = 0, tapT = null;
  const inRect = (x, y, r, pad) => r && x >= r.x - pad && x <= r.x + (r.w ?? r.width) + pad && y >= r.y - pad && y <= r.y + (r.h ?? r.height) + pad;
  document.addEventListener('pointerup', e => {
    const hitHouse = inRect(e.clientX, e.clientY, bldg3.screenRect, 16);
    const hitLabel = inRect(e.clientX, e.clientY, bldg3.tag.getBoundingClientRect(), 18);
    if (!hitHouse && !hitLabel) return;
    taps++; clearTimeout(tapT);
    tapT = setTimeout(() => { taps = 0; }, 1200);
    if (taps >= 3) { taps = 0; demolish(); }
  }, { passive: true });

  // label projection
  const v = new THREE.Vector3();
  function place(tag, x, y, z, opacity, offY = 0) {
    v.set(x, y, z).project(camera);
    const r = canvas.getBoundingClientRect();
    const sx = (v.x * 0.5 + 0.5) * r.width, sy = (-v.y * 0.5 + 0.5) * r.height + offY;
    tag.style.transform = `translate(${sx.toFixed(1)}px, ${sy.toFixed(1)}px) translate(-50%, -100%)`;
    tag.style.opacity = opacity.toFixed(2);
  }

  // ---------- post-processing: bloom + SMAA + filmic output ----------
  let composer = null, bloomPass = null, smaaPass = null;
  if (PP) {
    try {
      const w0 = sceneRoot.clientWidth || 1280, h0 = sceneRoot.clientHeight || 720;
      composer = new PP.EffectComposer(renderer);
      composer.addPass(new PP.RenderPass(scene, camera));
      bloomPass = new PP.UnrealBloomPass(new THREE.Vector2(w0, h0), isMobile ? 0.5 : 0.62, 0.55, 0.82);
      composer.addPass(bloomPass);
      if (!isMobile) { smaaPass = new PP.SMAAPass(w0, h0); composer.addPass(smaaPass); }
      composer.addPass(new PP.OutputPass());
    } catch (e) { console.warn('composer setup failed:', e); composer = null; }
  }

  const camBase = { px: 0, py: 12.5, pz: 23.5, lx: 0, ly: 0.3, lz: 4.2 };
  function resize() {
    const w = sceneRoot.clientWidth, h = sceneRoot.clientHeight;
    const pr = renderer.getPixelRatio();
    if (canvas.width !== Math.round(w * pr) || canvas.height !== Math.round(h * pr)) {
      renderer.setSize(w, h, false);
      camera.aspect = w / h;
      // portrait: look DOWN the corridor (pipeline recedes to the horizon) — fits a tall screen
      if (camera.aspect < 1.4) {
        camera.fov = 50;
        Object.assign(camBase, { px: -23, py: 10, pz: 4.6, lx: 12, ly: -2.2, lz: 1.4 });
      } else {
        camera.fov = 44;
        Object.assign(camBase, { px: 0, py: 12.5, pz: 23.5, lx: 0, ly: 0.3, lz: 4.2 });
      }
      camera.updateProjectionMatrix();
      if (composer) composer.setSize(w, h);
      if (bloomPass) bloomPass.setSize(w, h);
      if (smaaPass) smaaPass.setSize(w, h);
    }
  }

  let active = true;
  new IntersectionObserver(en => { active = en[0].isIntersecting; }, { threshold: 0 }).observe(canvas);

  renderer.setAnimationLoop(() => {
    if (!active || window.__pauseGL) return;
    resize();
    const t = clock.getElapsedTime();

    // drone along the corridor
    const k = (t % LOOP) / LOOP;
    const dx = X0 + (X1 - X0) * k;
    const alt = 5.4 + Math.sin(t * 1.6) * 0.18;
    drone.position.set(dx, alt, 0.6);
    drone.rotation.z = -0.06 + Math.sin(t * 1.1) * 0.03;
    drone.rotation.x = Math.sin(t * 0.9) * 0.02;
    for (const r of drone.rotors) r.rotation.y = t * 40;
    const blink = Math.sin(t * 7) > 0;
    drone.navL.material.opacity = blink ? 0.9 : 0.15;
    drone.navR.material.opacity = blink ? 0.15 : 0.9;
    cone.position.set(dx, alt - coneH / 2 - 0.25, 0.6);
    cone.material.opacity = 0.10 + Math.sin(t * 4.2) * 0.035;
    spot.position.x = dx; spot.position.z = 0.6;
    scanRing.position.x = dx; scanRing.position.z = 0.6;
    scanRing.material.opacity = 0.45 + Math.sin(t * 4.2) * 0.2;
    const sr = 1 + Math.sin(t * 4.2) * 0.06; scanRing.scale.set(sr, sr, sr);

    // detection boxes stay locked on every house; brighter glow as the drone scans past
    for (const s of settlements) {
      const d = Math.abs(dx - s.x);
      const scan = THREE.MathUtils.clamp(1 - (d - 2.4) / 1.8, 0, 1);
      const crushed = s === bldg3 && demo;
      s.det.material.opacity = crushed ? 0 : 0.42 + scan * 0.45 + Math.sin(t * 5) * 0.05;
      place(s.tag, s.labelPos.x, s.labelPos.y, s.labelPos.z, crushed && !s.tagLock ? 0 : 1, camera.aspect < 1.4 ? (s.mobOff || 0) : 0);
      if (s === bldg3) {
        v.set(s.x, 1.0, s.z).project(camera);
        const r = canvas.getBoundingClientRect();
        s.screenRect = { x: (v.x * .5 + .5) * r.width - 70 + r.left, y: (-v.y * .5 + .5) * r.height - 60 + r.top, w: 140, h: 120 };
      }
    }

    // radar rings
    rings.forEach((r, i) => {
      const ph = ((t * 0.45 + i * 0.5) % 1);
      r.scale.setScalar(0.3 + ph * 3.4);
      r.material.opacity = 0.5 * (1 - ph);
    });

    // satellite drift
    const st = (t % 60) / 60;
    sat.position.set(50 - st * 100, 26, -38);
    sat.rotation.z = 0.15;

    // camera idle + parallax
    parX += (tgtParX - parX) * 0.05; parY += (tgtParY - parY) * 0.05;
    camera.position.x = camBase.px + parX + Math.sin(t * 0.07) * 0.9;
    camera.position.y = camBase.py - parY + Math.sin(t * 0.11) * 0.25;
    camera.position.z = camBase.pz;
    camera.lookAt(camBase.lx, camBase.ly, camBase.lz);

    // demolition timeline
    if (demo) {
      const e = t - demo.t0;
      const { driveIn, push, hold, out } = DEMO;
      const bx = bldg3.x;
      if (!bldg3.tagLock && e >= driveIn + push * 0.9) {
        bldg3.tagLock = true;
        bldg3.tag.classList.add('tag3d-cleared');
        bldg3.tag.innerHTML = `BLDG-00${bldg3.id} &middot; ${bldg3.dist} — CLEARED`;
      }
      if (e < driveIn) {
        dozer.position.x = 46 + (bx + 2.6 - 46) * THREE.MathUtils.smoothstep(e / driveIn, 0, 1);
      } else if (e < driveIn + push) {
        const p = (e - driveIn) / push;
        dozer.position.x = bx + 2.6 - p * 2.2;
        const c = THREE.MathUtils.clamp(p * 1.4, 0, 1);
        bldg3.hutGroup.scale.y = 1 - c * 0.88;
        bldg3.hutGroup.rotation.z = -c * 0.3;
        bldg3.hutGroup.position.x = -c * 0.8;
        dusts.forEach((d, i) => {
          d.material.opacity = Math.max(0, 0.55 - p * 0.4 - i * 0.04);
          d.position.set(bx - 1 + Math.sin(i * 2.4) * 1.6, 0.4 + p * 2 + i * 0.18, bldg3.z + Math.cos(i * 1.7));
          d.scale.setScalar(1.2 + p * 2.4 + i * 0.2);
        });

      } else if (e < driveIn + push + hold) {
        bldg3.hutGroup.scale.y = 0.12; bldg3.hutGroup.rotation.z = -0.3; bldg3.hutGroup.position.x = -0.8;
        dusts.forEach(d => d.material.opacity *= 0.96);
      } else if (e < driveIn + push + hold + out) {
        const p = (e - driveIn - push - hold) / out;
        dozer.position.x = bx + 0.4 + p * (48 - bx);
        dusts.forEach(d => d.material.opacity *= 0.94);
      } else if (e < driveIn + push + hold + out + DEMO.regrow) {
        dozer.visible = false;
        const p = (e - driveIn - push - hold - out) / DEMO.regrow;
        bldg3.hutGroup.scale.y = 0.12 + p * 0.88;
        bldg3.hutGroup.rotation.z = -0.3 * (1 - p);
        bldg3.hutGroup.position.x = -0.8 * (1 - p);
      } else {
        bldg3.hutGroup.scale.y = 1; bldg3.hutGroup.rotation.z = 0; bldg3.hutGroup.position.x = 0;
        bldg3.tagLock = false;
        bldg3.tag.classList.remove('tag3d-cleared');
        bldg3.tag.innerHTML = `<i class="warn">!</i>BLDG-00${bldg3.id} &middot; ${bldg3.dist}`;
        dozer.position.x = 46;
        demo = null; running = false;
      }
    }

    if (composer) {
      try { composer.render(); }
      catch (e) { console.warn('post-processing render failed, falling back:', e); composer = null; }
    } else {
      renderer.render(scene, camera);
    }
  });
}
