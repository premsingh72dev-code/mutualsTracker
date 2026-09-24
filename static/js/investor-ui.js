/* Presentation and accessibility only. No data fetching, persistence, or calculations. */
(function () {
  'use strict';
  document.addEventListener('DOMContentLoaded', () => {
    const workspace = document.getElementById('dashboard-main');
    const tabs = Array.from(document.querySelectorAll('.tab-button'));
    const title = document.getElementById('workspace-title');
    const description = document.getElementById('workspace-description');
    const views = {
      'tab-best-funds': ['Find the funds worth your attention.', 'A focused view of relative strength, risk, and return consistency.'],
      'tab-overview': ['Explore the possibilities.', 'Search your fund universe, refine your criteria, and build your basket.'],
      'tab-rolling': ['Keep the longer view in focus.', 'Compare return patterns across different holding periods.'],
      'tab-charts': ['Put performance in perspective.', 'See how your fund universe compares across risk, scale, and categories.'],
      'tab-dynamic': ['Know the data behind the decision.', 'Review the source columns that power your research.']
    };
    function updateView() {
      const hasData = workspace && workspace.style.display !== 'none';
      tabs.forEach(tab => {
        tab.disabled = !hasData;
        if (tab.classList.contains('active')) tab.setAttribute('aria-current', 'page');
        else tab.removeAttribute('aria-current');
        tab.setAttribute('aria-controls', tab.dataset.tab);
      });
      const active = tabs.find(tab => tab.classList.contains('active'));
      const copy = hasData && active ? views[active.dataset.tab] : ['Make room for a clearer view.', 'Turn your fund data into a focused investment perspective.'];
      if (copy && title.textContent !== copy[0]) title.textContent = copy[0];
      if (copy && description.textContent !== copy[1]) description.textContent = copy[1];
    }
    updateView();
    if (workspace) new MutationObserver(updateView).observe(workspace, { attributes: true, attributeFilter: ['style'] });
    tabs.forEach(tab => new MutationObserver(updateView).observe(tab, { attributes: true, attributeFilter: ['class'] }));

    // Escape and focus restoration for existing dialogs, using existing close controls.
    const dialogs = document.querySelectorAll('.modal-backdrop');
    dialogs.forEach(backdrop => {
      let returnFocus = null;
      let open = false;
      const dialog = backdrop.querySelector('[role="dialog"]');
      const focusable = () => Array.from(backdrop.querySelectorAll('button, input, select, a[href], [tabindex="0"]')).filter(el => !el.disabled && el.getClientRects().length);
      new MutationObserver(() => {
        const nextOpen = backdrop.style.display !== 'none';
        if (nextOpen && !open) {
          returnFocus = document.activeElement;
          const field = backdrop.querySelector('input') || focusable()[0];
          if (field) field.focus();
        } else if (!nextOpen && open && returnFocus && returnFocus.isConnected) {
          returnFocus.focus();
        }
        open = nextOpen;
      }).observe(backdrop, { attributes: true, attributeFilter: ['style'] });
      backdrop.addEventListener('keydown', event => {
        if (event.key === 'Escape') {
          const close = backdrop.querySelector('.btn-modal-close');
          if (close) { event.preventDefault(); close.click(); }
        }
        if (event.key === 'Tab') {
          const nodes = focusable();
          if (!nodes.length) return;
          const first = nodes[0], last = nodes[nodes.length - 1];
          if (event.shiftKey && document.activeElement === first) { event.preventDefault(); last.focus(); }
          else if (!event.shiftKey && document.activeElement === last) { event.preventDefault(); first.focus(); }
        }
      });
      if (dialog) dialog.setAttribute('tabindex', '-1');
    });
  });
})();
