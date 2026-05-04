/* ════════════════════════════════════════════════
   CURSOR
════════════════════════════════════════════════ */
const cursor = document.getElementById('cursor');
const follower = document.getElementById('cursorFollower');
let mouseX = 0, mouseY = 0, follX = 0, follY = 0;

document.addEventListener('mousemove', e => {
  mouseX = e.clientX; mouseY = e.clientY;
  cursor.style.left = mouseX + 'px';
  cursor.style.top  = mouseY + 'px';
});

function animateFollower() {
  follX += (mouseX - follX) * 0.12;
  follY += (mouseY - follY) * 0.12;
  follower.style.left = follX + 'px';
  follower.style.top  = follY + 'px';
  requestAnimationFrame(animateFollower);
}
animateFollower();

document.querySelectorAll('a, button, .comp-card, .project-card, .skill-category, .tech-icon-card').forEach(el => {
  el.addEventListener('mouseenter', () => follower.classList.add('expanded'));
  el.addEventListener('mouseleave', () => follower.classList.remove('expanded'));
});

/* ════════════════════════════════════════════════
   PARTICLE CANVAS
════════════════════════════════════════════════ */
const canvas = document.getElementById('particleCanvas');
const ctx    = canvas.getContext('2d');

function resize() {
  canvas.width  = window.innerWidth;
  canvas.height = window.innerHeight;
}
resize();
window.addEventListener('resize', resize);

class Particle {
  constructor() { this.reset(); }
  reset() {
    this.x  = Math.random() * canvas.width;
    this.y  = Math.random() * canvas.height;
    this.vx = (Math.random() - .5) * .3;
    this.vy = (Math.random() - .5) * .3;
    this.r  = Math.random() * 1.5 + .3;
    this.a  = Math.random() * .5 + .1;
    this.hue = Math.random() < .5 ? 260 : 190; // purple or cyan
  }
  update() {
    this.x += this.vx; this.y += this.vy;
    if (this.x < 0 || this.x > canvas.width || this.y < 0 || this.y > canvas.height) this.reset();
  }
  draw() {
    ctx.save();
    ctx.globalAlpha = this.a;
    ctx.fillStyle = `hsl(${this.hue}, 70%, 65%)`;
    ctx.beginPath();
    ctx.arc(this.x, this.y, this.r, 0, Math.PI * 2);
    ctx.fill();
    ctx.restore();
  }
}

const PARTICLES = 90;
const particles = Array.from({ length: PARTICLES }, () => new Particle());

// Connection lines
function drawConnections() {
  const maxDist = 120;
  for (let i = 0; i < particles.length; i++) {
    for (let j = i + 1; j < particles.length; j++) {
      const dx = particles[i].x - particles[j].x;
      const dy = particles[i].y - particles[j].y;
      const dist = Math.sqrt(dx * dx + dy * dy);
      if (dist < maxDist) {
        ctx.save();
        ctx.globalAlpha = (1 - dist / maxDist) * .12;
        ctx.strokeStyle = '#8b5cf6';
        ctx.lineWidth = .5;
        ctx.beginPath();
        ctx.moveTo(particles[i].x, particles[i].y);
        ctx.lineTo(particles[j].x, particles[j].y);
        ctx.stroke();
        ctx.restore();
      }
    }
  }
}

function animateParticles() {
  ctx.clearRect(0, 0, canvas.width, canvas.height);
  particles.forEach(p => { p.update(); p.draw(); });
  drawConnections();
  requestAnimationFrame(animateParticles);
}
animateParticles();

/* ════════════════════════════════════════════════
   NAVBAR SCROLL
════════════════════════════════════════════════ */
const navbar = document.getElementById('navbar');
const sections = document.querySelectorAll('section[id]');

window.addEventListener('scroll', () => {
  if (window.scrollY > 40) navbar.classList.add('scrolled');
  else navbar.classList.remove('scrolled');

  // Active nav link
  const scrollY = window.scrollY + 120;
  sections.forEach(sec => {
    const top = sec.offsetTop;
    const h   = sec.offsetHeight;
    const id  = sec.getAttribute('id');
    const link = document.querySelector(`.nav-link[href="#${id}"]`);
    if (link) {
      if (scrollY >= top && scrollY < top + h) link.classList.add('active');
      else link.classList.remove('active');
    }
  });
});

/* ════════════════════════════════════════════════
   HAMBURGER
════════════════════════════════════════════════ */
const hamburger = document.getElementById('hamburger');
const navLinks  = document.querySelector('.nav-links');
hamburger.addEventListener('click', () => {
  navLinks.classList.toggle('open');
});

/* ════════════════════════════════════════════════
   SMOOTH SCROLL
════════════════════════════════════════════════ */
function scrollToSection(id) {
  const el = document.getElementById(id);
  if (el) el.scrollIntoView({ behavior: 'smooth' });
}
document.querySelectorAll('a[href^="#"]').forEach(a => {
  a.addEventListener('click', e => {
    e.preventDefault();
    const id = a.getAttribute('href').slice(1);
    scrollToSection(id);
    navLinks.classList.remove('open');
  });
});

/* ════════════════════════════════════════════════
   INTERSECTION OBSERVER — REVEAL
════════════════════════════════════════════════ */
const revealObs = new IntersectionObserver((entries) => {
  entries.forEach(entry => {
    if (entry.isIntersecting) {
      entry.target.classList.add('visible');
    }
  });
}, { threshold: .12 });

document.querySelectorAll('.reveal-up, .reveal-left, .reveal-right').forEach(el => revealObs.observe(el));

/* ════════════════════════════════════════════════
   SKILL BARS
════════════════════════════════════════════════ */
const skillObs = new IntersectionObserver((entries) => {
  entries.forEach(entry => {
    if (entry.isIntersecting) {
      entry.target.querySelectorAll('.skill-fill').forEach(bar => {
        const w = bar.dataset.width;
        setTimeout(() => { bar.style.width = w + '%'; }, 200);
      });
      skillObs.unobserve(entry.target);
    }
  });
}, { threshold: .3 });

document.querySelectorAll('.skill-category').forEach(el => skillObs.observe(el));

/* ════════════════════════════════════════════════
   COUNTER ANIMATION
════════════════════════════════════════════════ */
const counterObs = new IntersectionObserver((entries) => {
  entries.forEach(entry => {
    if (entry.isIntersecting) {
      entry.target.querySelectorAll('.stat-number').forEach(el => {
        const target = +el.dataset.target;
        let current = 0;
        const step = target / 40;
        const timer = setInterval(() => {
          current += step;
          if (current >= target) { el.textContent = target; clearInterval(timer); }
          else el.textContent = Math.floor(current);
        }, 40);
      });
      counterObs.unobserve(entry.target);
    }
  });
}, { threshold: .5 });

document.querySelectorAll('.about-stats').forEach(el => counterObs.observe(el));

/* ════════════════════════════════════════════════
   TERMINAL TYPEWRITER
════════════════════════════════════════════════ */
const sequences = [
  {
    cmd: 'nmap -sV 192.168.1.0/24',
    output: [
      { text: 'Starting Nmap scan...', cls: 'out-info' },
      { text: 'Host: 192.168.1.1  [open]', cls: 'out-line' },
      { text: '80/tcp  http   Apache 2.4', cls: 'out-line' },
      { text: '443/tcp https  OpenSSL', cls: 'out-line' },
      { text: 'Scan complete. 3 hosts up.', cls: 'out-info' },
    ]
  },
  {
    cmd: 'glpi --version && systemctl status glpi',
    output: [
      { text: 'GLPI v10.0.12 — Running', cls: 'out-line' },
      { text: '● glpi.service — Active (running)', cls: 'out-line' },
      { text: 'Uptime: 14 days, 3 hours', cls: 'out-info' },
    ]
  },
  {
    cmd: 'az vm list --output table',
    output: [
      { text: 'Name         Location    Status', cls: 'out-info' },
      { text: 'web-server   francecentral  Running', cls: 'out-line' },
      { text: 'glpi-prod    francecentral  Running', cls: 'out-line' },
    ]
  },
];

const typedCmd  = document.getElementById('typedCmd');
const termOut   = document.getElementById('termOutput');
const caret     = document.getElementById('caret');
let seqIdx = 0;

function sleep(ms) { return new Promise(r => setTimeout(r, ms)); }

async function runSequence() {
  const seq = sequences[seqIdx % sequences.length];
  seqIdx++;

  // Clear
  typedCmd.textContent = '';
  termOut.innerHTML = '';
  caret.style.display = 'inline';

  // Type command
  for (const ch of seq.cmd) {
    typedCmd.textContent += ch;
    await sleep(55);
  }

  caret.style.display = 'none';
  await sleep(400);

  // Show output lines one by one
  for (const line of seq.output) {
    await sleep(300);
    const div = document.createElement('div');
    div.className = line.cls;
    div.textContent = '  ' + line.text;
    div.style.opacity = '0';
    div.style.transform = 'translateX(-8px)';
    div.style.transition = 'opacity .3s, transform .3s';
    termOut.appendChild(div);
    requestAnimationFrame(() => {
      div.style.opacity = '1';
      div.style.transform = 'translateX(0)';
    });
  }

  await sleep(3000);
  runSequence();
}

// Start after a short delay
setTimeout(runSequence, 800);

/* ════════════════════════════════════════════════
   FORM SUBMIT
════════════════════════════════════════════════ */
document.getElementById('contactForm').addEventListener('submit', function(e) {
  e.preventDefault();
  const btn = this.querySelector('button[type="submit"]');
  const orig = btn.innerHTML;
  btn.innerHTML = '✓ Message envoyé !';
  btn.style.background = 'linear-gradient(135deg, #059669, #10b981)';
  setTimeout(() => {
    btn.innerHTML = orig;
    btn.style.background = '';
    this.reset();
  }, 3000);
});

/* ════════════════════════════════════════════════
   GLOWING BORDER on cards — mouse tracking
════════════════════════════════════════════════ */
document.querySelectorAll('.comp-card, .project-card, .skill-category, .profile-card, .contact-card, .contact-form').forEach(card => {
  card.addEventListener('mousemove', e => {
    const rect = card.getBoundingClientRect();
    const x = ((e.clientX - rect.left) / rect.width * 100).toFixed(1);
    const y = ((e.clientY - rect.top)  / rect.height * 100).toFixed(1);
    card.style.background = `radial-gradient(circle at ${x}% ${y}%, rgba(139,92,246,.06) 0%, rgba(255,255,255,.03) 60%)`;
  });
  card.addEventListener('mouseleave', () => {
    card.style.background = '';
  });
});

/* ════════════════════════════════════════════════
   PAGE LOAD FADE-IN
════════════════════════════════════════════════ */
document.body.style.opacity = '0';
document.body.style.transition = 'opacity .6s';
window.addEventListener('load', () => {
  document.body.style.opacity = '1';
});
