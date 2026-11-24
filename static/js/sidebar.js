/**
 * App sidebar — folds into an icon rail on desktop, slides in as a sheet on ≤1024px.
 * Behaviour follows the public description of Arc's sidebar-rail block; the code is our own.
 *
 * Desktop
 *   - Header button or "[" (outside text fields, menus and dialogs) folds/unfolds the rail.
 *     The choice is kept in localStorage ("sidebar-collapsed") and restored before first paint.
 *   - Folded: each icon shows its label in a tooltip; a group opens a flyout with its pages.
 *   - Unfolded: groups expand and collapse by height.
 * ≤1024px
 *   - The header button opens a sheet (focus moves to the current page, focus stays inside).
 *   - Escape, the scrim, the close button or a drag to the left closes it.
 *   - Widening past 1024px closes it.
 */
(function () {
    'use strict';

    const MOBILE = window.matchMedia('(max-width: 1024px)');
    const REDUCED = window.matchMedia('(prefers-reduced-motion: reduce)');
    const html = document.documentElement;
    const STORE_KEY = 'sidebar-collapsed';

    function init() {
        const sidebar = document.getElementById('app-sidebar');
        const main = document.querySelector('.main-content');
        const toggle = document.getElementById('sidebar-toggle');
        const overlay = document.getElementById('sidebar-overlay');
        const closeBtn = document.getElementById('sidebar-close-btn');
        if (!sidebar) return;

        const isRail = () => html.classList.contains('sb-rail') && !MOBILE.matches;

        /* ── Current page ── */
        sidebar.querySelectorAll('.nav-item.active:not(.submenu-toggle), .submenu-item.active')
            .forEach((a) => a.setAttribute('aria-current', 'page'));

        /* ── Rail ── */
        function syncToggle() {
            if (!toggle) return;
            let label;
            if (MOBILE.matches) {
                const open = sidebar.classList.contains('mobile-open');
                label = open ? 'Close menu' : 'Open menu';
                toggle.setAttribute('aria-expanded', String(open));
            } else {
                const rail = html.classList.contains('sb-rail');
                label = rail ? 'Expand sidebar' : 'Collapse sidebar';
                toggle.setAttribute('aria-expanded', String(!rail));
            }
            toggle.setAttribute('aria-label', label);
            toggle.title = MOBILE.matches ? label : `${label}  [`;
        }

        function setRail(on) {
            hideTip(true);
            closeFlyout(false);
            html.classList.toggle('sb-rail', on);
            try { localStorage.setItem(STORE_KEY, String(on)); } catch (_) { /* storage blocked */ }
            syncToggle();
            // charts measure their container: tell them once the width has settled
            const done = () => window.dispatchEvent(new Event('resize'));
            if (REDUCED.matches) { done(); return; }
            let fired = false;
            const onEnd = (e) => {
                if (e.target !== sidebar || e.propertyName !== 'width' || fired) return;
                fired = true;
                sidebar.removeEventListener('transitionend', onEnd);
                done();
            };
            sidebar.addEventListener('transitionend', onEnd);
            setTimeout(() => { if (!fired) { fired = true; sidebar.removeEventListener('transitionend', onEnd); done(); } }, 400);
        }

        /* ── Sheet (≤1024px) ── */
        let returnFocus = null;

        function openSheet() {
            if (sidebar.classList.contains('mobile-open')) return;
            returnFocus = document.activeElement;
            sidebar.classList.add('mobile-open');
            sidebar.setAttribute('role', 'dialog');
            sidebar.setAttribute('aria-modal', 'true');
            if (overlay) overlay.classList.add('active');
            if (main) main.inert = true;
            document.documentElement.style.overflow = 'hidden';
            syncToggle();
            // focus the current page so keyboard and screen reader users land in place
            const current = sidebar.querySelector('[aria-current="page"]') ||
                sidebar.querySelector('.submenu-toggle.active') || sidebar.querySelector('.nav-item');
            requestAnimationFrame(() => {
                if (current) {
                    current.focus({ preventScroll: true });
                    current.scrollIntoView({ block: 'center' });
                }
            });
        }

        function closeSheet(restoreFocus = true) {
            if (!sidebar.classList.contains('mobile-open')) return;
            sidebar.classList.remove('mobile-open');
            sidebar.removeAttribute('role');
            sidebar.removeAttribute('aria-modal');
            sidebar.style.removeProperty('--sb-drag');
            if (overlay) { overlay.classList.remove('active'); overlay.style.removeProperty('--sb-scrim'); }
            if (main) main.inert = false;
            document.documentElement.style.overflow = '';
            syncToggle();
            if (restoreFocus && returnFocus && returnFocus !== document.body && document.contains(returnFocus)) returnFocus.focus({ preventScroll: true });
            else if (restoreFocus && toggle) toggle.focus({ preventScroll: true });
        }

        if (toggle) {
            toggle.addEventListener('click', () => {
                if (MOBILE.matches) {
                    sidebar.classList.contains('mobile-open') ? closeSheet() : openSheet();
                } else {
                    setRail(!html.classList.contains('sb-rail'));
                }
            });
        }
        if (overlay) overlay.addEventListener('click', () => closeSheet());
        if (closeBtn) closeBtn.addEventListener('click', () => closeSheet());

        // keep focus inside the open sheet
        sidebar.addEventListener('keydown', (e) => {
            if (!sidebar.classList.contains('mobile-open')) return;
            if (e.key === 'Escape') { e.preventDefault(); closeSheet(); return; }
            if (e.key !== 'Tab') return;
            const focusables = Array.from(sidebar.querySelectorAll('a[href], button:not([disabled])'))
                .filter((el) => el.offsetParent !== null);
            if (!focusables.length) return;
            const first = focusables[0];
            const last = focusables[focusables.length - 1];
            if (e.shiftKey && document.activeElement === first) { e.preventDefault(); last.focus(); }
            else if (!e.shiftKey && document.activeElement === last) { e.preventDefault(); first.focus(); }
        });

        MOBILE.addEventListener('change', () => {
            closeSheet(false);
            hideTip(true);
            closeFlyout(false);
            syncToggle();
        });

        /* ── Drag the sheet to the left to close it ── */
        (function bindSheetDrag() {
            let startX = 0, startY = 0, lastX = 0, lastT = 0, v = 0, mode = null, id = null;
            const width = () => sidebar.offsetWidth || 288;

            sidebar.addEventListener('pointerdown', (e) => {
                if (!sidebar.classList.contains('mobile-open') || e.pointerType === 'mouse' || REDUCED.matches) return;
                startX = lastX = e.clientX; startY = e.clientY;
                lastT = performance.now(); v = 0; mode = 'pending'; id = e.pointerId;
            });
            sidebar.addEventListener('pointermove', (e) => {
                if (!mode || e.pointerId !== id) return;
                const dx = e.clientX - startX;
                const dy = e.clientY - startY;
                if (mode === 'pending') {
                    if (Math.abs(dx) < 8 && Math.abs(dy) < 8) return;
                    if (Math.abs(dy) > Math.abs(dx) || dx > 0) { mode = null; return; }   // scrolling, or pulling right
                    mode = 'drag';
                    sidebar.setPointerCapture(id);
                    sidebar.setAttribute('data-dragging', '');
                }
                const now = performance.now();
                v = (e.clientX - lastX) / Math.max(1, now - lastT);
                lastX = e.clientX; lastT = now;
                const x = Math.min(0, dx);
                sidebar.style.setProperty('--sb-drag', x + 'px');
                if (overlay) overlay.style.setProperty('--sb-scrim', String(1 + x / width()));
            });
            const end = () => {
                if (mode !== 'drag') { mode = null; return; }
                mode = null;
                sidebar.removeAttribute('data-dragging');
                const dx = lastX - startX;
                if (dx < -80 || v < -0.5) {
                    closeSheet(false);
                } else {
                    sidebar.style.removeProperty('--sb-drag');
                    if (overlay) overlay.style.removeProperty('--sb-scrim');
                }
            };
            sidebar.addEventListener('pointerup', end);
            sidebar.addEventListener('pointercancel', end);
        })();

        /* ── Groups: height animation when unfolded, flyout when folded ── */
        function setGroup(group, open) {
            const btn = group.querySelector(':scope > .submenu-toggle');
            const panel = group.querySelector(':scope > .submenu');
            if (!btn || !panel) return;
            btn.setAttribute('aria-expanded', String(open));
            if (REDUCED.matches) { group.classList.toggle('open', open); return; }

            const from = panel.getBoundingClientRect().height;
            group.classList.toggle('open', open);
            panel.style.height = 'auto';
            const to = open ? panel.scrollHeight : 0;
            panel.style.height = from + 'px';
            void panel.offsetHeight;                // commit the start height
            panel.style.height = to + 'px';
            const onEnd = (e) => {
                if (e.target !== panel || e.propertyName !== 'height') return;
                panel.removeEventListener('transitionend', onEnd);
                panel.style.removeProperty('height');   // back to the CSS value (auto or 0)
            };
            panel.addEventListener('transitionend', onEnd);
        }

        sidebar.querySelectorAll('.has-submenu').forEach((group) => {
            const btn = group.querySelector(':scope > .submenu-toggle');
            const panel = group.querySelector(':scope > .submenu');
            if (!btn || !panel) return;
            if (panel.id === '') panel.id = 'sb-group-' + Math.random().toString(36).slice(2, 8);
            btn.setAttribute('aria-controls', panel.id);
            btn.setAttribute('aria-expanded', String(group.classList.contains('open')));
            btn.addEventListener('click', (e) => {
                e.preventDefault();
                if (isRail()) {
                    flyoutFor === group ? closeFlyout(true) : openFlyout(group, e.detail === 0);
                } else {
                    setGroup(group, !group.classList.contains('open'));
                }
            });
        });

        /* ── Rail tooltip ── */
        const tip = document.createElement('div');
        tip.className = 'sb-tip';
        tip.setAttribute('role', 'tooltip');
        tip.id = 'sb-tip';
        document.body.appendChild(tip);
        let tipTimer = null;
        let tipWarmUntil = 0;     // moving between icons shows the next label at once
        let tipFor = null;

        function labelOf(item) {
            const l = item.querySelector('.label');
            return (l ? l.textContent : item.textContent).trim();
        }

        function showTip(item, instant) {
            if (!isRail() || flyoutFor) return;
            clearTimeout(tipTimer);
            const run = () => {
                tipFor = item;
                tip.textContent = labelOf(item);
                const r = item.getBoundingClientRect();
                tip.style.left = (sidebar.getBoundingClientRect().right + 8) + 'px';
                tip.style.top = (r.top + r.height / 2) + 'px';
                tip.setAttribute('data-show', '');
                item.setAttribute('aria-describedby', 'sb-tip');
            };
            if (instant || performance.now() < tipWarmUntil) run();
            else tipTimer = setTimeout(run, 350);
        }

        function hideTip(immediate) {
            clearTimeout(tipTimer);
            if (tipFor) tipFor.removeAttribute('aria-describedby');
            if (tip.hasAttribute('data-show')) tipWarmUntil = performance.now() + 400;
            tipFor = null;
            tip.removeAttribute('data-show');
            if (immediate) tipWarmUntil = 0;
        }

        sidebar.querySelectorAll('.sidebar-nav > .nav-item, .has-submenu > .submenu-toggle').forEach((item) => {
            item.addEventListener('pointerenter', (e) => { if (e.pointerType === 'mouse') showTip(item); });
            item.addEventListener('pointerleave', () => hideTip());
            item.addEventListener('focus', () => { if (item.matches(':focus-visible')) showTip(item, true); });
            item.addEventListener('blur', () => hideTip());
        });
        sidebar.querySelector('.sidebar-nav')?.addEventListener('scroll', () => hideTip(true), { passive: true });

        /* ── Rail flyout for groups ── */
        const flyout = document.createElement('div');
        flyout.className = 'sb-flyout';
        flyout.hidden = true;
        document.body.appendChild(flyout);
        let flyoutFor = null;
        let flyoutTimer = null;

        function openFlyout(group, viaKeyboard) {
            hideTip(true);
            clearTimeout(flyoutTimer);
            const btn = group.querySelector(':scope > .submenu-toggle');
            const links = Array.from(group.querySelectorAll(':scope > .submenu > .submenu-item'));
            flyout.innerHTML = '';
            const title = document.createElement('div');
            title.className = 'sb-flyout-title';
            title.textContent = labelOf(btn);
            title.id = 'sb-flyout-title';
            flyout.appendChild(title);
            flyout.setAttribute('role', 'menu');
            flyout.setAttribute('aria-labelledby', 'sb-flyout-title');
            links.forEach((a) => {
                const c = document.createElement('a');
                c.href = a.getAttribute('href');
                c.textContent = a.textContent.trim();
                c.setAttribute('role', 'menuitem');
                c.tabIndex = -1;
                if (a.getAttribute('aria-current')) c.setAttribute('aria-current', 'page');
                flyout.appendChild(c);
            });
            if (!links.length) {
                const empty = document.createElement('div');
                empty.className = 'sb-flyout-title';
                empty.textContent = 'No sensors';
                flyout.appendChild(empty);
            }

            flyoutFor = group;
            btn.setAttribute('aria-haspopup', 'menu');
            flyout.hidden = false;
            const r = btn.getBoundingClientRect();
            const left = sidebar.getBoundingClientRect().right + 8;
            const h = flyout.offsetHeight;
            const top = Math.max(12, Math.min(r.top - 6, window.innerHeight - h - 12));
            flyout.style.left = left + 'px';
            flyout.style.top = top + 'px';
            flyout.style.transformOrigin = `0 ${r.top + r.height / 2 - top}px`;
            requestAnimationFrame(() => flyout.setAttribute('data-show', ''));

            const items = flyout.querySelectorAll('a');
            if (viaKeyboard && items.length) {
                const cur = flyout.querySelector('[aria-current]') || items[0];
                cur.focus({ preventScroll: true });
            }
            document.addEventListener('pointerdown', onOutside, true);
        }

        function closeFlyout(restoreFocus) {
            if (!flyoutFor) return;
            const btn = flyoutFor.querySelector(':scope > .submenu-toggle');
            flyoutFor = null;
            flyout.removeAttribute('data-show');
            flyoutTimer = setTimeout(() => { flyout.hidden = true; }, REDUCED.matches ? 0 : 130);
            document.removeEventListener('pointerdown', onOutside, true);
            if (restoreFocus && btn) btn.focus({ preventScroll: true });
        }

        function onOutside(e) {
            if (flyout.contains(e.target)) return;
            if (flyoutFor && flyoutFor.contains(e.target)) return;   // its own button toggles
            closeFlyout(false);
        }

        flyout.addEventListener('keydown', (e) => {
            const items = Array.from(flyout.querySelectorAll('a'));
            const i = items.indexOf(document.activeElement);
            if (e.key === 'ArrowDown') { e.preventDefault(); items[(i + 1) % items.length]?.focus(); }
            else if (e.key === 'ArrowUp') { e.preventDefault(); items[(i - 1 + items.length) % items.length]?.focus(); }
            else if (e.key === 'Home') { e.preventDefault(); items[0]?.focus(); }
            else if (e.key === 'End') { e.preventDefault(); items[items.length - 1]?.focus(); }
            else if (e.key === 'Escape' || e.key === 'ArrowLeft') { e.preventDefault(); closeFlyout(true); }
            else if (e.key === 'Tab') { closeFlyout(false); }
        });
        window.addEventListener('resize', () => { hideTip(true); if (flyoutFor) closeFlyout(false); });

        /* ── "[" folds the rail ── */
        document.addEventListener('keydown', (e) => {
            const bracket = e.key === '[' || (e.code === 'BracketLeft' && !e.shiftKey);
            if (!bracket || e.ctrlKey || e.metaKey || e.altKey || e.repeat || MOBILE.matches) return;
            const t = e.target;
            if (t && (t.isContentEditable || /^(INPUT|TEXTAREA|SELECT)$/.test(t.tagName))) return;
            if (t && t.closest && t.closest('[role="dialog"], [role="menu"], [role="listbox"], .modal.show')) return;
            if (document.querySelector('.modal.show, .status-modal.active')) return;
            e.preventDefault();
            setRail(!html.classList.contains('sb-rail'));
        });

        syncToggle();
        // restored state painted without animation; enable transitions from now on
        requestAnimationFrame(() => requestAnimationFrame(() => html.classList.remove('sb-booting')));
    }

    if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init);
    else init();
})();
