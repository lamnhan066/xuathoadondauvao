const VI_MONTHS = [
    "Tháng 1",
    "Tháng 2",
    "Tháng 3",
    "Tháng 4",
    "Tháng 5",
    "Tháng 6",
    "Tháng 7",
    "Tháng 8",
    "Tháng 9",
    "Tháng 10",
    "Tháng 11",
    "Tháng 12"
];

const VI_WEEKDAYS = ["T2", "T3", "T4", "T5", "T6", "T7", "CN"];

function createDatePicker() {
    const root = document.createElement("div");
    root.className = "custom-datepicker";
    root.style.display = "none";

    root.innerHTML = `
    <div class="dp-header">
      <div class="dp-nav">
        <button type="button" data-action="prev">‹</button>
      </div>
      <div class="dp-title">
        <select class="dp-month" aria-label="Chọn tháng"></select>
        <select class="dp-year" aria-label="Chọn năm"></select>
      </div>
      <div class="dp-nav">
        <button type="button" data-action="next">›</button>
      </div>
    </div>
    <div class="dp-body">
      <table>
        <thead><tr>${VI_WEEKDAYS.map(d => `<th>${d}</th>`).join("")}</tr></thead>
        <tbody></tbody>
      </table>
    </div>
    <div class="dp-footer dp-footer">
      <button type="button" data-action="today">Hôm nay</button>
      <button type="button" data-action="close">Đóng</button>
    </div>
  `;

    document.body.appendChild(root);

    const titleEl = root.querySelector('.dp-title');
    const monthSelect = root.querySelector('.dp-month');
    const yearSelect = root.querySelector('.dp-year');
    const tbody = root.querySelector('tbody');

    let attachedInput = null;
    let viewDate = new Date();

    function formatInputDate(d) {
        const y = d.getFullYear();
        const m = String(d.getMonth() + 1).padStart(2, '0');
        const day = String(d.getDate()).padStart(2, '0');
        return `${y}-${m}-${day}`;
    }

    function render() {
        const year = viewDate.getFullYear();
        const month = viewDate.getMonth();

        // populate month select
        if (monthSelect) {
            monthSelect.innerHTML = VI_MONTHS.map((m, i) => `<option value="${i}">${m}</option>`).join('');
            monthSelect.value = String(month);
        }

        // populate year select (ensure selected year is included)
        if (yearSelect) {
            const nowYear = new Date().getFullYear();
            const start = Math.min(viewDate.getFullYear() - 5, nowYear - 10);
            const end = Math.max(viewDate.getFullYear() + 5, nowYear + 10);
            let years = [];
            for (let y = start; y <= end; y++) years.push(y);
            yearSelect.innerHTML = years.map((y) => `<option value="${y}">${y}</option>`).join('');
            yearSelect.value = String(year);
        }

        // first day of month (0 Sun .. 6 Sat)
        const first = new Date(year, month, 1);
        const last = new Date(year, month + 1, 0);
        const daysInMonth = last.getDate();

        // Monday-first calendar index
        const firstWeekday = (first.getDay() + 6) % 7; // 0..6

        const cells = [];
        for (let i = 0; i < firstWeekday; i++) cells.push('');
        for (let d = 1; d <= daysInMonth; d++) cells.push(d);
        while (cells.length % 7 !== 0) cells.push('');

        const selectedVal = attachedInput && attachedInput.value ? attachedInput.value : null;

        let html = '';
        for (let r = 0; r < cells.length; r += 7) {
            html += '<tr>';
            for (let c = 0; c < 7; c++) {
                const v = cells[r + c];
                if (!v) {
                    html += '<td></td>';
                    continue;
                }
                const dateObj = new Date(year, month, v);
                const classes = ['day'];
                const today = new Date();
                if (dateObj.toDateString() === today.toDateString()) classes.push('today');
                if (selectedVal) {
                    const sel = new Date(selectedVal + 'T00:00:00');
                    if (dateObj.toDateString() === sel.toDateString()) classes.push('selected');
                }
                html += `<td><div class="${classes.join(' ')}" data-day="${v}">${v}</div></td>`;
            }
            html += '</tr>';
        }

        tbody.innerHTML = html;
    }

    function show(input) {
        attachedInput = input;
        if (input.value) {
            const parts = input.value.split('-');
            if (parts.length === 3) {
                viewDate = new Date(Number(parts[0]), Number(parts[1]) - 1, 1);
            }
        }
        render();
        root.style.display = '';
        positionNearInput(input, root);
        setTimeout(() => { document.addEventListener('mousedown', onDocClick); document.addEventListener('keydown', onKey); }, 0);
    }

    function hide() {
        root.style.display = 'none';
        attachedInput = null;
        document.removeEventListener('mousedown', onDocClick);
        document.removeEventListener('keydown', onKey);
    }

    function onDocClick(e) {
        if (!root.contains(e.target) && e.target !== attachedInput) hide();
    }

    function onKey(e) {
        if (e.key === 'Escape') hide();
    }

    function positionNearInput(input, elNode) {
        const rect = input.getBoundingClientRect();
        let top = rect.bottom + window.scrollY + 6;
        let left = rect.left + window.scrollX;
        if (left + elNode.offsetWidth > window.innerWidth - 8) {
            left = Math.max(8, window.innerWidth - elNode.offsetWidth - 8);
        }
        if (top + elNode.offsetHeight > window.innerHeight - 8) {
            top = rect.top + window.scrollY - elNode.offsetHeight - 6;
        }
        elNode.style.top = `${top}px`;
        elNode.style.left = `${left}px`;
    }

    root.addEventListener('click', (e) => {
        const action = e.target.closest('[data-action]')?.dataset?.action;
        if (action === 'prev') {
            viewDate.setMonth(viewDate.getMonth() - 1);
            render();
            return;
        }
        if (action === 'next') {
            viewDate.setMonth(viewDate.getMonth() + 1);
            render();
            return;
        }
        if (action === 'today') {
            const now = new Date();
            if (attachedInput) {
                attachedInput.value = formatInputDate(now);
                attachedInput.dispatchEvent(new Event('change', { bubbles: true }));
            }
            hide();
            return;
        }
        if (action === 'close') {
            hide();
            return;
        }

        const dayEl = e.target.closest('.day');
        if (dayEl && attachedInput) {
            const day = Number(dayEl.dataset.day);
            const sel = new Date(viewDate.getFullYear(), viewDate.getMonth(), day);
            attachedInput.value = formatInputDate(sel);
            attachedInput.dispatchEvent(new Event('change', { bubbles: true }));
            hide();
        }
    });

    // wire month/year selects to change the view
    if (monthSelect) {
        monthSelect.addEventListener('change', () => {
            viewDate.setMonth(Number(monthSelect.value));
            render();
        });
    }

    if (yearSelect) {
        yearSelect.addEventListener('change', () => {
            viewDate.setFullYear(Number(yearSelect.value));
            render();
        });
    }

    return { show, hide };
}

export function initDatePicker(inputFrom, inputTo) {
    const datePicker = createDatePicker();
    if (!inputFrom && !inputTo) return datePicker;

    ['focus', 'click', 'touchend'].forEach(evt => {
        if (inputFrom) inputFrom.addEventListener(evt, (e) => { e.preventDefault(); datePicker.show(inputFrom); });
        if (inputTo) inputTo.addEventListener(evt, (e) => { e.preventDefault(); datePicker.show(inputTo); });
    });

    return datePicker;
}
