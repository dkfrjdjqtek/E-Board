// 2026.06.09 Changed: 모든 문서 화면에서 공통 사용하는 날짜 DateBox/LocalDateKey 헬퍼로 정리
(function (window) {
    'use strict';

    function unwrap(container) {
        return container && container.jquery ? container[0] : container;
    }

    function readJsonScript(id, fallbackValue) {
        const el = document.getElementById(id);
        if (!el) return fallbackValue;

        try {
            const raw = el.textContent || '';
            if (!raw.trim()) return fallbackValue;
            return JSON.parse(raw);
        } catch (e) {
            console.error('JSON parse failed:', id, e);
            return fallbackValue;
        }
    }

    function pad2(value) {
        const n = parseInt(String(value ?? '0'), 10);
        return n < 10 ? '0' + n : String(n);
    }

    function normalizeCultureName(value) {
        const s = String(value || '').trim().replace('_', '-');
        return s || 'ko-KR';
    }

    function getLocalDatePattern(cultureName) {
        const culture = normalizeCultureName(cultureName).toLowerCase();

        if (culture === 'vi-vn') return 'dd/MM/yyyy';
        if (culture === 'en-us') return 'MM/dd/yyyy';
        if (culture === 'id-id') return 'dd/MM/yyyy';
        if (culture === 'zh-cn') return 'yyyy/MM/dd';

        return 'yyyy-MM-dd';
    }

    function getLocalDateTimePattern(cultureName) {
        return getLocalDatePattern(cultureName) + ' HH:mm';
    }

    function textOrFallback(value, resourceKey, fallbackValue) {
        const s = String(value || '').trim();
        return !s || s === resourceKey ? fallbackValue : s;
    }

    function lowerFirst(value) {
        const s = String(value || '');
        if (!s) return s;
        return s.charAt(0).toLowerCase() + s.slice(1);
    }

    function upperFirst(value) {
        const s = String(value || '');
        if (!s) return s;
        return s.charAt(0).toUpperCase() + s.slice(1);
    }

    function uniqueArray(values) {
        const result = [];

        (values || []).forEach(value => {
            const s = String(value || '').trim();
            if (!s) return;
            if (!result.includes(s)) result.push(s);
        });

        return result;
    }

    function makeFieldCandidates(field) {
        const s = String(field || '').trim();
        if (!s) return [];

        return uniqueArray([
            s,
            lowerFirst(s),
            upperFirst(s)
        ]);
    }

    function getFieldValue(row, fields, fallbackValue) {
        if (!row) return fallbackValue;

        const sourceFields = Array.isArray(fields) ? fields : [fields];
        const candidates = [];

        sourceFields.forEach(field => {
            makeFieldCandidates(field).forEach(candidate => candidates.push(candidate));
        });

        const uniqueCandidates = uniqueArray(candidates);

        for (const field of uniqueCandidates) {
            if (Object.prototype.hasOwnProperty.call(row, field)) {
                const value = row[field];
                if (value !== null && value !== undefined && String(value).trim() !== '') {
                    return value;
                }
            }
        }

        return fallbackValue;
    }

    function parseDateTextToKey(value, datePattern) {
        const s = String(value || '').trim();
        if (!s) return '';

        let m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
        if (m) {
            return m[1] + '-' + m[2] + '-' + m[3];
        }

        m = s.match(/^(\d{4})\/(\d{2})\/(\d{2})/);
        if (m) {
            return m[1] + '-' + m[2] + '-' + m[3];
        }

        m = s.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
        if (m) {
            const pattern = String(datePattern || '').toLowerCase();

            if (pattern.indexOf('mm/dd') === 0) {
                return m[3] + '-' + m[1] + '-' + m[2];
            }

            return m[3] + '-' + m[2] + '-' + m[1];
        }

        return '';
    }

    function dateBoxValueToDateKey(value, datePattern) {
        if (value === null || value === undefined || value === '') {
            return '';
        }

        if (value instanceof Date && !isNaN(value.getTime())) {
            return value.getFullYear()
                + '-' + pad2(value.getMonth() + 1)
                + '-' + pad2(value.getDate());
        }

        const parsedTextKey = parseDateTextToKey(value, datePattern);
        if (parsedTextKey) {
            return parsedTextKey;
        }

        const d = new Date(String(value));
        if (!isNaN(d.getTime())) {
            return d.getFullYear()
                + '-' + pad2(d.getMonth() + 1)
                + '-' + pad2(d.getDate());
        }

        return '';
    }

    function createDateEditorOptions(text, options) {
        const sourceText = text || {};
        const sourceOptions = options || {};
        const displayFormat =
            sourceOptions.displayFormat ||
            sourceText.datePattern ||
            getLocalDatePattern(sourceText.currentCulture || sourceOptions.currentCulture || 'ko-KR');

        return {
            type: 'date',
            pickerType: 'calendar',
            displayFormat: displayFormat,
            showDropDownButton: true,
            showClearButton: sourceOptions.showClearButton === true,
            applyValueMode: 'useButtons',
            applyButtonText: textOrFallback(sourceText.ok, '_CM_OK', '확인'),
            cancelButtonText: textOrFallback(sourceText.cancel, '_CM_Cancel', '취소'),
            todayButtonText: textOrFallback(sourceText.today, '_CM_Today', '금일')
        };
    }

    function createLocalDateFilterExpression(localDateKeyField, filterValue, datePattern) {
        const key = dateBoxValueToDateKey(filterValue, datePattern);
        if (!key) return null;

        return [localDateKeyField, '=', key];
    }

    function renderLocalTextCell(container, data, localTextField, emptyText) {
        const el = unwrap(container);
        if (!el) return;

        el.textContent = getFieldValue(data, localTextField, emptyText || '');
    }

    function createLocalDateColumn(options) {
        const sourceOptions = options || {};
        const text = sourceOptions.text || {};
        const localTextField = sourceOptions.localTextField;
        const localDateKeyField = sourceOptions.localDateKeyField;
        const sortFields = sourceOptions.sortFields || sourceOptions.sortField || localDateKeyField;
        const datePattern =
            sourceOptions.datePattern ||
            text.datePattern ||
            getLocalDatePattern(text.currentCulture || sourceOptions.currentCulture || 'ko-KR');

        if (!localTextField) {
            throw new Error('DocControllerHelper.createLocalDateColumn: localTextField is required.');
        }

        if (!localDateKeyField) {
            throw new Error('DocControllerHelper.createLocalDateColumn: localDateKeyField is required.');
        }

        const column = {
            caption: sourceOptions.caption || '',
            dataField: localDateKeyField,
            dataType: 'date',
            width: sourceOptions.width || 150,
            alignment: sourceOptions.alignment || 'center',
            cssClass: sourceOptions.cssClass || '',
            allowHeaderFiltering: sourceOptions.allowHeaderFiltering === true,
            allowFiltering: sourceOptions.allowFiltering !== false,
            allowSorting: sourceOptions.allowSorting !== false,
            editorOptions: createDateEditorOptions(text, {
                displayFormat: datePattern,
                showClearButton: sourceOptions.showClearButton === true
            }),
            calculateSortValue: function (row) {
                return getFieldValue(
                    row,
                    Array.isArray(sortFields)
                        ? sortFields.concat([localDateKeyField, localTextField])
                        : [sortFields, localDateKeyField, localTextField],
                    ''
                );
            },
            calculateFilterExpression: function (filterValue) {
                return createLocalDateFilterExpression(localDateKeyField, filterValue, datePattern);
            },
            cellTemplate: function (container, cellOptions) {
                renderLocalTextCell(
                    container,
                    cellOptions ? cellOptions.data : null,
                    localTextField,
                    text.empty || ''
                );
            }
        };

        if (sourceOptions.name) column.name = sourceOptions.name;
        if (sourceOptions.fixed !== undefined) column.fixed = sourceOptions.fixed;
        if (sourceOptions.fixedPosition !== undefined) column.fixedPosition = sourceOptions.fixedPosition;
        if (sourceOptions.minWidth !== undefined) column.minWidth = sourceOptions.minWidth;
        if (sourceOptions.visible !== undefined) column.visible = sourceOptions.visible;
        if (sourceOptions.visibleIndex !== undefined) column.visibleIndex = sourceOptions.visibleIndex;

        return Object.assign(column, sourceOptions.columnOptions || {});
    }

    function normalizeLocalDateRow(row, options) {
        if (!row) return row;

        const sourceOptions = options || {};
        const localTextField = sourceOptions.localTextField;
        const localDateKeyField = sourceOptions.localDateKeyField;

        if (localTextField) {
            const textValue = getFieldValue(row, localTextField, '');
            makeFieldCandidates(localTextField).forEach(field => {
                row[field] = textValue;
            });
        }

        if (localDateKeyField) {
            const keyValue = getFieldValue(row, localDateKeyField, '');
            makeFieldCandidates(localDateKeyField).forEach(field => {
                row[field] = keyValue;
            });
        }

        return row;
    }

    function normalizeLocalDateRows(rows, options) {
        if (!Array.isArray(rows)) return rows;

        rows.forEach(row => {
            normalizeLocalDateRow(row, options);
        });

        return rows;
    }

    function setElementText(element, value, emptyText) {
        const el = unwrap(element);
        if (!el) return;

        const s = value === null || value === undefined || String(value).trim() === ''
            ? (emptyText || '')
            : String(value);

        el.textContent = s;
    }

    function renderLocalDateText(element, row, localTextField, emptyText) {
        setElementText(
            element,
            getFieldValue(row, localTextField, emptyText || ''),
            emptyText || ''
        );
    }

    window.DocControllerHelper = Object.assign(window.DocControllerHelper || {}, {
        unwrap: unwrap,
        readJsonScript: readJsonScript,
        pad2: pad2,
        normalizeCultureName: normalizeCultureName,
        getLocalDatePattern: getLocalDatePattern,
        getLocalDateTimePattern: getLocalDateTimePattern,
        textOrFallback: textOrFallback,
        getFieldValue: getFieldValue,
        dateBoxValueToDateKey: dateBoxValueToDateKey,
        createDateEditorOptions: createDateEditorOptions,
        createLocalDateFilterExpression: createLocalDateFilterExpression,
        createLocalDateColumn: createLocalDateColumn,
        normalizeLocalDateRow: normalizeLocalDateRow,
        normalizeLocalDateRows: normalizeLocalDateRows,
        setElementText: setElementText,
        renderLocalDateText: renderLocalDateText
    });
})(window);