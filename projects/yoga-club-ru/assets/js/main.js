/* ============================================================
   САДХАНА — интерактив и анимации
   Принципы:
   1. Интерфейс (меню, аккордеон, фильтры) работает БЕЗ GSAP.
   2. Анимации включаются, только если GSAP загрузился
      и пользователь не просил уменьшить движение.
   3. Любая ошибка анимации → .motion-off → страница полностью видна.
   ============================================================ */
(function () {
  'use strict';

  var root = document.documentElement;
  var reduceMQ = window.matchMedia('(prefers-reduced-motion: reduce)');
  var hasGSAP = typeof window.gsap !== 'undefined' &&
                typeof window.ScrollTrigger !== 'undefined';
  var motionOK = hasGSAP && !reduceMQ.matches;

  if (!motionOK) root.classList.add('motion-off');

  var $  = function (s, ctx) { return (ctx || document).querySelector(s); };
  var $$ = function (s, ctx) { return Array.prototype.slice.call((ctx || document).querySelectorAll(s)); };

  /* ---------- Текущий год в подвале ---------- */
  var yearEl = $('#year');
  if (yearEl) yearEl.textContent = new Date().getFullYear();

  /* =========================================================
     1. МОБИЛЬНОЕ МЕНЮ
     ========================================================= */
  (function drawer() {
    var burger = $('#burger');
    var panel  = $('#drawer');
    if (!burger || !panel) return;

    var open = false;

    function setOpen(state) {
      open = state;
      burger.setAttribute('aria-expanded', String(state));
      burger.setAttribute('aria-label', state ? 'Закрыть меню' : 'Открыть меню');
      panel.classList.toggle('is-open', state);
      panel.setAttribute('aria-hidden', String(!state));
      document.body.classList.toggle('is-locked', state);
      if (state) {
        var first = panel.querySelector('a');
        if (first) first.focus();
      }
    }

    burger.addEventListener('click', function () { setOpen(!open); });

    panel.addEventListener('click', function (e) {
      if (e.target.closest('a')) setOpen(false);
    });

    document.addEventListener('keydown', function (e) {
      if (e.key === 'Escape' && open) { setOpen(false); burger.focus(); }
    });

    window.addEventListener('resize', function () {
      if (open && window.innerWidth > 864) setOpen(false);
    });
  }());

  /* =========================================================
     2. АККОРДЕОН (FAQ) — работает без GSAP, через height + CSS transition
     ========================================================= */
  (function accordion() {
    var wrap = $('[data-accordion]');
    if (!wrap) return;

    var triggers = $$('.accordion__trigger', wrap);

    function close(btn) {
      var panel = document.getElementById(btn.getAttribute('aria-controls'));
      if (!panel) return;
      panel.style.height = panel.scrollHeight + 'px';
      void panel.offsetHeight;               /* форсируем reflow */
      panel.style.height = '0px';
      btn.setAttribute('aria-expanded', 'false');
    }

    function open(btn) {
      var panel = document.getElementById(btn.getAttribute('aria-controls'));
      if (!panel) return;
      panel.style.height = panel.scrollHeight + 'px';
      btn.setAttribute('aria-expanded', 'true');
      panel.addEventListener('transitionend', function once(e) {
        if (e.propertyName !== 'height') return;
        panel.removeEventListener('transitionend', once);
        if (btn.getAttribute('aria-expanded') === 'true') panel.style.height = 'auto';
      });
    }

    triggers.forEach(function (btn) {
      btn.addEventListener('click', function () {
        var isOpen = btn.getAttribute('aria-expanded') === 'true';
        /* «гармошка»: одновременно открыт один вопрос */
        triggers.forEach(function (other) {
          if (other !== btn && other.getAttribute('aria-expanded') === 'true') close(other);
        });
        isOpen ? close(btn) : open(btn);
      });
    });

    /* Пересчёт высоты открытого блока при смене ширины экрана */
    var resizeTimer;
    window.addEventListener('resize', function () {
      clearTimeout(resizeTimer);
      resizeTimer = setTimeout(function () {
        triggers.forEach(function (btn) {
          if (btn.getAttribute('aria-expanded') !== 'true') return;
          var panel = document.getElementById(btn.getAttribute('aria-controls'));
          if (panel) panel.style.height = 'auto';
        });
      }, 150);
    });
  }());

  /* =========================================================
     3. ФИЛЬТР РАСПИСАНИЯ (вкладки)
     ========================================================= */
  (function tabs() {
    var wrap = $('[data-tabs]');
    var week = $('#week');
    if (!wrap || !week) return;

    var list      = $('.tabs__list', wrap);
    var indicator = $('.tabs__indicator', wrap);
    var tabEls    = $$('.tabs__tab', wrap);

    function moveIndicator(tab) {
      if (!indicator || !list) return;
      indicator.style.width = tab.offsetWidth + 'px';
      indicator.style.transform = 'translateX(' + (tab.offsetLeft - list.clientLeft) + 'px)';
    }

    function applyFilter(filter) {
      $$('.day', week).forEach(function (day) {
        var slots = $$('.day__slot', day);
        var visible = 0;
        slots.forEach(function (slot) {
          var track = slot.getAttribute('data-track');
          var show = filter === 'all' || track === filter || track === 'rest';
          slot.hidden = !show;
          if (show) visible++;
        });
        var empty = $('.day__empty', day);
        if (empty) empty.hidden = visible !== 0;
      });
    }

    function select(tab) {
      tabEls.forEach(function (t) {
        var active = t === tab;
        t.setAttribute('aria-selected', String(active));
        t.tabIndex = active ? 0 : -1;
      });
      week.setAttribute('aria-labelledby', tab.id);
      moveIndicator(tab);
      applyFilter(tab.getAttribute('data-filter'));
    }

    tabEls.forEach(function (tab) {
      tab.addEventListener('click', function () { select(tab); });
    });

    /* Клавиатура: стрелки, Home, End — паттерн ARIA tablist */
    list.addEventListener('keydown', function (e) {
      var i = tabEls.indexOf(document.activeElement);
      if (i < 0) return;
      var next = null;
      if (e.key === 'ArrowRight') next = tabEls[(i + 1) % tabEls.length];
      else if (e.key === 'ArrowLeft') next = tabEls[(i - 1 + tabEls.length) % tabEls.length];
      else if (e.key === 'Home') next = tabEls[0];
      else if (e.key === 'End') next = tabEls[tabEls.length - 1];
      if (!next) return;
      e.preventDefault();
      next.focus();
      select(next);
    });

    var active = tabEls.filter(function (t) { return t.getAttribute('aria-selected') === 'true'; })[0] || tabEls[0];
    /* Ждём загрузки шрифтов: ширина вкладок до этого другая */
    if (document.fonts && document.fonts.ready) {
      document.fonts.ready.then(function () { moveIndicator(active); });
    }
    requestAnimationFrame(function () { moveIndicator(active); });
    window.addEventListener('resize', function () { moveIndicator(
      tabEls.filter(function (t) { return t.getAttribute('aria-selected') === 'true'; })[0] || tabEls[0]
    ); });
  }());

  /* =========================================================
     4. КАРУСЕЛЬ ОТЗЫВОВ
     ========================================================= */
  (function quotes() {
    var track = $('[data-quotes]');
    var nav   = $('[data-quotes-nav]');
    if (!track || !nav) return;

    function step(dir) {
      var card = track.querySelector('.quote');
      if (!card) return;
      var gap = parseFloat(getComputedStyle(track).columnGap || '24') || 24;
      track.scrollBy({
        left: dir * (card.getBoundingClientRect().width + gap),
        behavior: reduceMQ.matches ? 'auto' : 'smooth'
      });
    }

    $$('button', nav).forEach(function (btn) {
      btn.addEventListener('click', function () { step(parseInt(btn.getAttribute('data-dir'), 10)); });
    });

    function syncDisabled() {
      var maxScroll = track.scrollWidth - track.clientWidth - 2;
      $$('button', nav).forEach(function (btn) {
        var dir = parseInt(btn.getAttribute('data-dir'), 10);
        var atEdge = dir < 0 ? track.scrollLeft <= 2 : track.scrollLeft >= maxScroll;
        btn.disabled = atEdge;
        btn.style.opacity = atEdge ? '.4' : '1';
      });
    }
    track.addEventListener('scroll', syncDisabled, { passive: true });
    window.addEventListener('resize', syncDisabled);
    syncDisabled();
  }());


  /* =========================================================
     4b. ЭКРАН ЗАГРУЗКИ — снимается при любом исходе
     ========================================================= */
  var LOADER_ANIM = 1.15;   /* заставка должна быть короткой: ~1.4 сек до конца */
  (function loader() {
    var el = $('#loader');
    if (!el) return;

    function hide(instant) {
      if (el.dataset.done) return;
      el.dataset.done = '1';
      if (instant || !motionOK) { el.hidden = true; return; }
      gsap.to(el, {
        opacity: 0, duration: 0.5, ease: 'power2.inOut',
        onComplete: function () { el.hidden = true; }
      });
    }

    if (!motionOK) { hide(true); return; }

    var fill = $('#loaderFill'), num = $('#loaderNum'), mark = $('#loaderMark');
    var shapes = $$('path, circle', mark);
    shapes.forEach(function (sh) {
      var len = 0;
      try { len = sh.getTotalLength(); } catch (e) { len = 0; }
      if (len) gsap.set(sh, { strokeDasharray: len, strokeDashoffset: len });
    });

    var tl = gsap.timeline();
    tl.to(shapes, { strokeDashoffset: 0, duration: 0.65, stagger: 0.06, ease: 'power2.inOut' }, 0);
    if (fill) tl.to(fill, { scaleX: 1, duration: 0.8, ease: 'power1.inOut' }, 0);
    if (num) {
      var o = { v: 0 };
      tl.to(o, {
        v: 100, duration: 0.8, ease: 'power1.inOut',
        onUpdate: function () { num.textContent = Math.round(o.v); }
      }, 0);
    }
    tl.to(el, { yPercent: -100, duration: 0.55, ease: 'expo.inOut' }, '+=0.05')
      .add(function () { el.hidden = true; });

    /* страховка: заставка не должна залипнуть ни при каких обстоятельствах */
    setTimeout(function () { hide(false); }, 2600);
  }());

  /* =========================================================
     5. АНИМАЦИИ (только при наличии GSAP и без reduced-motion)
     ========================================================= */
  if (!motionOK) return;

  try {
    gsap.registerPlugin(ScrollTrigger);
    gsap.defaults({ ease: 'power2.out' });

    var header = $('#header');

    /* --- 5.1 Шапка: фон при прокрутке + прогресс чтения --- */
    ScrollTrigger.create({
      start: 'top -80',
      end: 99999,
      onUpdate: function (self) {
        if (header) header.classList.toggle('is-stuck', self.scroll() > 80);
      }
    });

    gsap.to('.scroll-progress', {
      scaleX: 1,
      ease: 'none',
      scrollTrigger: { start: 0, end: 'max', scrub: 0.3 }
    });

    /* --- 5.2 Вступительная анимация первого экрана --- */
    var intro = gsap.timeline({ delay: $('#loader') ? LOADER_ANIM + 0.35 : 0.15 });

    intro.to('.hero .line-mask > span', {
      y: 0, duration: 1.05, stagger: 0.09, ease: 'expo.out'
    });

    var heroReveals = $$('.hero [data-reveal]');
    if (heroReveals.length) {
      intro.to(heroReveals, {
        opacity: 1, y: 0, scale: 1, duration: 0.8, stagger: 0.1
      }, '-=0.75');
    }

    /* Счётчики в первом экране */
    $$('[data-count]').forEach(function (el) {
      var target = parseInt(el.getAttribute('data-count'), 10);
      if (isNaN(target)) return;
      var obj = { v: 0 };
      intro.to(obj, {
        v: target, duration: 1.1, ease: 'power1.out',
        onUpdate: function () { el.textContent = Math.round(obj.v); }
      }, '-=0.6');
    });

    /* --- 5.3 Помощники видимости ---
       ВАЖНО: браузер восстанавливает позицию прокрутки при обновлении
       страницы, а ссылка вида /#schedule открывает её сразу с середины.
       Всё, что к моменту запуска уже выше линии старта, нужно показать
       немедленно — иначе целые секции останутся невидимыми.            */
    var START_RATIO = 0.88;

    function rectOf(el) { return el.getBoundingClientRect(); }
    function isAboveView(el) { return rectOf(el).bottom < 0; }
    function isPastStart(el) { return rectOf(el).top < window.innerHeight * START_RATIO; }

    /* --- 5.4 Заголовки секций: построчное раскрытие --- */
    $$('.line-mask').forEach(function (mask) {
      if (mask.closest('.hero')) return;
      var line = mask.firstElementChild;
      if (!line) return;

      if (isAboveView(mask)) { gsap.set(line, { y: 0 }); return; }

      var opts = { y: 0, duration: 0.95, ease: 'expo.out' };
      if (isPastStart(mask)) { gsap.to(line, opts); return; }

      opts.scrollTrigger = { trigger: mask, start: 'top 88%', once: true };
      gsap.to(line, opts);
    });

    /* --- 5.5 Появление блоков при прокрутке --- */
    var reveals = $$('[data-reveal]').filter(function (el) { return !el.closest('.hero'); });

    var shownInstantly = [], shownNow = [], pending = [];
    reveals.forEach(function (el) {
      if (isAboveView(el)) shownInstantly.push(el);
      else if (isPastStart(el)) shownNow.push(el);
      else pending.push(el);
    });

    if (shownInstantly.length) {
      gsap.set(shownInstantly, { opacity: 1, y: 0, scale: 1, clearProps: 'transform' });
    }
    if (shownNow.length) {
      gsap.to(shownNow, {
        opacity: 1, y: 0, scale: 1, duration: 0.7,
        /* amount ограничивает ВЕСЬ каскад: сколько бы блоков ни попало
           в первый экран, они проявятся максимум за 0.5 сек */
        stagger: { each: 0.06, amount: 0.5 },
        clearProps: 'transform'
      });
    }
    if (pending.length) {
      ScrollTrigger.batch(pending, {
        start: 'top 88%',
        once: true,
        onEnter: function (batch) {
          gsap.to(batch, {
            opacity: 1, y: 0, scale: 1,
            duration: 0.7, stagger: { each: 0.09, amount: 0.45 },
            clearProps: 'transform'
          });
        }
      });
    }

    /* --- 5.6 Прорисовка линейных иконок --- */
    $$('[data-draw]').forEach(function (svg) {
      var shapes = $$('path, circle, line', svg);
      var lengths = shapes.map(function (shape) {
        var len = 0;
        try { len = shape.getTotalLength(); } catch (err) { len = 0; }
        return len;
      });

      function prime() {
        shapes.forEach(function (shape, i) {
          if (!lengths[i]) { gsap.set(shape, { opacity: 1 }); return; }
          gsap.set(shape, { strokeDasharray: lengths[i], strokeDashoffset: lengths[i], opacity: 1 });
        });
      }
      function draw(instant) {
        gsap.to(shapes, {
          strokeDashoffset: 0,
          duration: instant ? 0 : 1.1,
          stagger: instant ? 0 : 0.09,
          ease: 'power2.inOut',
          onComplete: function () {
            gsap.set(shapes, { clearProps: 'strokeDasharray,strokeDashoffset' });
          }
        });
      }

      if (isAboveView(svg)) { gsap.set(shapes, { opacity: 1 }); return; }

      prime();
      if (isPastStart(svg)) { draw(false); return; }

      ScrollTrigger.create({
        trigger: svg, start: 'top 88%', once: true,
        onEnter: function () { draw(false); }
      });
    });

    /* --- 5.7 Параллакс декоративных пятен (только фон, не текст) --- */
    $$('[data-parallax]').forEach(function (layer) {
      var depth = parseFloat(layer.getAttribute('data-parallax')) || 0.08;
      gsap.to(layer, {
        yPercent: depth * 100,
        ease: 'none',
        scrollTrigger: { trigger: '.hero', start: 'top top', end: 'bottom top', scrub: 0.6 }
      });
    });

    /* --- 5.8 Круг дыхания: 4 сек вдох / 6 сек выдох --- */
    (function breath() {
      var core  = $('#breathCore');
      var label = $('#breathLabel');
      var wrap  = $('#breath');
      if (!core || !wrap) return;

      var tl = gsap.timeline({ repeat: -1 });
      tl.to(core, {
        scale: 1.28, duration: 4, ease: 'sine.inOut',
        onStart: function () { if (label) label.textContent = 'вдох'; }
      })
      .to(core, {
        scale: 1, duration: 6, ease: 'sine.inOut',
        onStart: function () { if (label) label.textContent = 'выдох'; }
      });

      /* Не тратим кадры, когда блок вне экрана или вкладка скрыта */
      ScrollTrigger.create({
        trigger: wrap,
        start: 'top bottom', end: 'bottom top',
        onToggle: function (self) { self.isActive ? tl.play() : tl.pause(); }
      });
      document.addEventListener('visibilitychange', function () {
        document.hidden ? tl.pause() : tl.play();
      });
    }());

    /* --- 5.9 Бегущая строка --- */
    (function marquee() {
      var wrap = $('[data-marquee]');
      if (!wrap) return;
      var track = $('.marquee__track', wrap);
      if (!track) return;

      /* Дублируем содержимое, чтобы цикл был бесшовным */
      var original = track.innerHTML;
      track.innerHTML = original + original;

      var half = track.scrollWidth / 2;
      var tl = gsap.to(track, {
        x: -half,
        duration: half / 45,          /* ~45 px/сек — спокойный темп */
        ease: 'none',
        repeat: -1
      });

      wrap.addEventListener('pointerenter', function () { tl.timeScale(0.25); });
      wrap.addEventListener('pointerleave', function () { tl.timeScale(1); });

      ScrollTrigger.create({
        trigger: wrap, start: 'top bottom', end: 'bottom top',
        onToggle: function (self) { self.isActive ? tl.play() : tl.pause(); }
      });
    }());

    /* --- 5.10 Вращающаяся печать в финальном блоке --- */
    (function seal() {
      var el = $('#seal');
      if (!el) return;
      var tl = gsap.to(el, { rotation: 360, duration: 44, ease: 'none', repeat: -1, transformOrigin: '50% 50%' });
      ScrollTrigger.create({
        trigger: el, start: 'top bottom', end: 'bottom top',
        onToggle: function (self) { self.isActive ? tl.play() : tl.pause(); }
      });
    }());

    /* --- 5.11 Точка-бегунок в подсказке прокрутки --- */
    gsap.fromTo('.scroll-cue__dot',
      { yPercent: -110 },
      { yPercent: 260, duration: 2.1, ease: 'power1.inOut', repeat: -1, repeatDelay: 0.4 }
    );

    /* --- 5.12 Подсветка активного пункта меню --- */
    $$('main section[id]').forEach(function (section) {
      var link = $('.nav__link[href="#' + section.id + '"]');
      if (!link) return;
      ScrollTrigger.create({
        trigger: section,
        start: 'top 45%',
        end: 'bottom 45%',
        onToggle: function (self) { link.classList.toggle('is-active', self.isActive); }
      });
    });

    /* --- 5.13 Тёмная секция: мягкое «проявление» фона --- */
    $$('.dark-section').forEach(function (section) {
      gsap.fromTo(section,
        { backgroundColor: '#2E3D30' },
        {
          backgroundColor: '#1F2A22',
          ease: 'none',
          scrollTrigger: { trigger: section, start: 'top bottom', end: 'top 40%', scrub: true }
        }
      );
    });


    /* --- 5.14 Параллакс внутри рамок изображений --- */
    $$('[data-parallax-img]').forEach(function (img) {
      var depth = parseFloat(img.getAttribute('data-parallax-img')) || 0.08;
      var host = img.closest('.frame, .hero__photo, .manifesto__bg, .section-photo') || img.parentElement;
      gsap.fromTo(img,
        { yPercent: -depth * 50 },
        {
          yPercent: depth * 50, ease: 'none',
          scrollTrigger: { trigger: host, start: 'top bottom', end: 'bottom top', scrub: 0.5 }
        });
    });

    /* --- 5.15 Шторка: изображение проявляется снизу вверх --- */
    $$('[data-curtain] img').forEach(function (img) {
      if (isAboveView(img)) { gsap.set(img, { clipPath: 'inset(0% 0 0 0)' }); return; }
      var tw = { clipPath: 'inset(0% 0 0 0)', duration: 1.15, ease: 'power3.out' };
      if (isPastStart(img)) { gsap.to(img, tw); return; }
      tw.scrollTrigger = { trigger: img, start: 'top 88%', once: true };
      gsap.to(img, tw);
    });

    /* --- 5.16 Манифест: слова загораются по мере прокрутки --- */
    (function manifesto() {
      var el = $('[data-words]');
      if (!el) return;
      var words = el.textContent.trim().split(/\s+/);
      el.textContent = '';
      words.forEach(function (w, i) {
        var sp = document.createElement('span');
        sp.className = 'w';
        sp.textContent = w;
        el.appendChild(sp);
        if (i < words.length - 1) el.appendChild(document.createTextNode(' '));
      });
      var spans = $$('.w', el);
      ScrollTrigger.create({
        trigger: el,
        start: 'top 78%',
        end: 'bottom 55%',
        scrub: true,
        onUpdate: function (self) {
          var n = Math.round(self.progress * spans.length);
          spans.forEach(function (sp, i) { sp.classList.toggle('on', i < n); });
        }
      });
    }());

    /* --- 5.17 Лента направлений: закрепляем и ведём вбок --- */
    (function rail() {
      var rail = $('[data-rail]'), track = $('[data-rail-track]');
      if (!rail || !track) return;
      var hint = $('[data-rail-hint] span');

      var mm = gsap.matchMedia();
      mm.add('(min-width: 62rem)', function () {
        rail.classList.add('is-pinned');
        var dist = function () { return Math.max(0, track.scrollWidth - window.innerWidth + 64); };
        var tw = gsap.to(track, {
          x: function () { return -dist(); },
          ease: 'none',
          scrollTrigger: {
            trigger: rail,
            pin: true,
            scrub: 0.6,
            start: 'top 14%',
            end: function () { return '+=' + dist(); },
            invalidateOnRefresh: true,
            onUpdate: function (self) {
              if (hint) gsap.set(hint, { scaleX: self.progress });
            }
          }
        });
        return function () {            /* откат при смене брейкпоинта */
          rail.classList.remove('is-pinned');
          tw.scrollTrigger && tw.scrollTrigger.kill();
          tw.kill();
          gsap.set(track, { clearProps: 'transform' });
        };
      });
    }());

    /* --- 5.18 Лента-галерея --- */
    (function strip() {
      var track = $('[data-strip]');
      if (!track) return;
      track.innerHTML = track.innerHTML + track.innerHTML;
      var half = track.scrollWidth / 2;
      var tl = gsap.to(track, { x: -half, duration: half / 38, ease: 'none', repeat: -1 });
      track.parentElement.addEventListener('pointerenter', function () { tl.timeScale(0.2); });
      track.parentElement.addEventListener('pointerleave', function () { tl.timeScale(1); });
      ScrollTrigger.create({
        trigger: track.parentElement, start: 'top bottom', end: 'bottom top',
        onToggle: function (self) { self.isActive ? tl.play() : tl.pause(); }
      });
    }());

    /* --- 5.19 Липкая стопка «три шага» --- */
    (function stack() {
      var cards = $$('[data-stack-card]');
      if (!cards.length) return;
      cards.forEach(function (card, i) {
        card.style.top = 'calc(7rem + ' + (i * 1.4) + 'rem)';
        if (i === cards.length - 1) return;
        gsap.to(card, {
          scale: 0.94, opacity: 0.55, ease: 'none',
          scrollTrigger: {
            trigger: cards[i + 1], start: 'top 80%', end: 'top 30%', scrub: true
          }
        });
      });
    }());

    /* --- 5.20 Курсор и магнитные кнопки --- */
    (function pointer() {
      if (!window.matchMedia('(hover: hover) and (pointer: fine)').matches) return;
      var ring = $('#cursor'), dot = $('#cursorDot');
      if (!ring || !dot) return;

      gsap.set([ring, dot], { xPercent: -50, yPercent: -50 });
      var rx = gsap.quickTo(ring, 'x', { duration: 0.45, ease: 'power3' });
      var ry = gsap.quickTo(ring, 'y', { duration: 0.45, ease: 'power3' });
      var dx = gsap.quickTo(dot, 'x', { duration: 0.12, ease: 'power3' });
      var dy = gsap.quickTo(dot, 'y', { duration: 0.12, ease: 'power3' });

      window.addEventListener('pointermove', function (e) {
        /* до первого движения мыши кольцо не должно висеть в углу экрана */
        if (!ring.classList.contains('is-live')) {
          gsap.set([ring, dot], { x: e.clientX, y: e.clientY });
          ring.classList.add('is-live'); dot.classList.add('is-live');
        }
        rx(e.clientX); ry(e.clientY); dx(e.clientX); dy(e.clientY);
        /* e.target уже указывает на элемент под курсором — hit-test не нужен */
        var over = e.target;
        var dark = !!(over && over.closest && over.closest('.dark-section'));
        ring.classList.toggle('is-over-dark', dark);
        dot.classList.toggle('is-over-dark', dark);
      }, { passive: true });

      $$('a, button, .frame, input, summary').forEach(function (el) {
        el.addEventListener('pointerenter', function () { gsap.to(ring, { scale: 1.7, duration: 0.3 }); });
        el.addEventListener('pointerleave', function () { gsap.to(ring, { scale: 1, duration: 0.3 }); });
      });

      /* магнит: кнопка чуть тянется к курсору */
      $$('.btn').forEach(function (btn) {
        var qx = gsap.quickTo(btn, 'x', { duration: 0.4, ease: 'power3' });
        var qy = gsap.quickTo(btn, 'y', { duration: 0.4, ease: 'power3' });
        btn.addEventListener('pointermove', function (e) {
          var r = btn.getBoundingClientRect();
          qx((e.clientX - (r.left + r.width / 2)) * 0.22);
          qy((e.clientY - (r.top + r.height / 2)) * 0.32);
        });
        btn.addEventListener('pointerleave', function () { qx(0); qy(0); });
      });
    }());

    ScrollTrigger.refresh();
    if (document.fonts && document.fonts.ready) {
      document.fonts.ready.then(function () { ScrollTrigger.refresh(); });
    }

  } catch (err) {
    /* Любая проблема с анимацией не должна ломать страницу */
    root.classList.add('motion-off');
    if (window.console && console.warn) console.warn('Анимации отключены:', err);
  }

  /* Страховка: если через 3 секунды блоки всё ещё невидимы — показываем их */
  setTimeout(function () {
    var probe = document.querySelector('.section [data-reveal]');
    if (!probe) return;
    if (parseFloat(getComputedStyle(probe).opacity) === 0 &&
        probe.getBoundingClientRect().top < window.innerHeight) {
      root.classList.add('motion-off');
    }
  }, 3000);

}());
