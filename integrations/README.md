# School Light Glass UI V2 (preview only)

This is an integration module for **Apps Sekolah**. The current `eschool/index.html` embeds a cross-origin Google Apps Script iframe. Do **not** add this file to the outer wrapper expecting access to controls inside the iframe: browser same-origin security prevents it.

## Integration in the ACTUAL Apps Script HTML

1. Add a client HTML file containing this module's JavaScript (without `<script>` tags), e.g. `SchoolGlassUI.html`.
2. Add `<script><?!= include('SchoolGlassUI'); ?></script>` near the end of the application HTML (only if its templating environment defines `include`; otherwise paste as an ordinary inline script).
3. Add `sg-glass` to the actual cards/panels and `sg-interactive` to relevant buttons.
4. On confirmed save success call `SchoolGlassUI.success('Berjaya disimpan')`; on confirmed failure call `SchoolGlassUI.error('Gagal menyimpan')`. Do NOT call these merely on tab opening.
5. To adjust settings call `SchoolGlassUI.setSettings({theme:'light',sound:true,volume:.75,motion:true,notifications:true})`. The module defaults to those settings and persists user changes to localStorage.
6. Test login, tab switching, keyboard focus, mobile viewport, and real backend response paths before deployment.

The module does not alter student records, backend calls or authentication. Does not globally force theme styling onto a host app; apply styles selectively after verifying existing selectors. No production deployment is performed by this preview PR.
