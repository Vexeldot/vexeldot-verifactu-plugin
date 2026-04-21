<?php
if (!defined('ABSPATH')) {
    exit;
}

class VVD_Admin {
    public static function init() {
        add_action('admin_menu', [__CLASS__, 'admin_menu']);
        add_action('admin_enqueue_scripts', [__CLASS__, 'enqueue']);
        add_action('admin_post_vvd_save_client', [__CLASS__, 'save_client']);
        add_action('admin_post_vvd_save_invoice', [__CLASS__, 'save_invoice']);
        add_action('admin_post_vvd_send_invoice_email', [__CLASS__, 'send_invoice_email']);
        add_action('admin_post_vvd_save_settings', [__CLASS__, 'save_settings']);
        add_action('admin_post_vvd_download_pdf', [__CLASS__, 'download_pdf']);
        add_action('admin_post_vvd_print_document', [__CLASS__, 'print_document']);
        add_action('admin_post_vvd_public_document', [__CLASS__, 'public_document']);
        add_action('admin_post_nopriv_vvd_public_document', [__CLASS__, 'public_document']);
        add_action('admin_post_vvd_delete_invoice', [__CLASS__, 'delete_invoice']);
        add_action('admin_post_vvd_delete_client', [__CLASS__, 'delete_client']);
        add_action('admin_post_vvd_convert_to_invoice', [__CLASS__, 'convert_to_invoice']);
        add_action('admin_post_vvd_import_accounting_csv', [__CLASS__, 'import_accounting_csv']);
        add_action('admin_post_vvd_import_legacy_xlsx', [__CLASS__, 'import_legacy_xlsx']);
        add_action('admin_post_vvd_toggle_invoice_paid', [__CLASS__, 'toggle_invoice_paid']);
        add_action('admin_post_vvd_bulk_update_invoices', [__CLASS__, 'bulk_update_invoices']);
        add_action('admin_post_vvd_import_clients_xlsx', [__CLASS__, 'import_clients_xlsx']);
        add_action('admin_post_vvd_delete_legacy_imports', [__CLASS__, 'delete_legacy_imports']);
        add_action('admin_post_vvd_delete_imported_clients', [__CLASS__, 'delete_imported_clients']);
        add_action('admin_post_vvd_prepare_invoice_email', [__CLASS__, 'prepare_invoice_email']);
    }

    protected static function table($name) {
        global $wpdb;
        return $wpdb->prefix . $name;
    }

    protected static function settings() {
        global $wpdb;
        return (array) $wpdb->get_row('SELECT * FROM ' . self::table('vvd_settings') . ' ORDER BY id ASC LIMIT 1', ARRAY_A);
    }

    protected static function series_rows() {
        global $wpdb;
        return $wpdb->get_results('SELECT * FROM ' . self::table('vvd_series') . ' ORDER BY id ASC', ARRAY_A);
    }

    protected static function client($id) {
        global $wpdb;
        return (array) $wpdb->get_row($wpdb->prepare('SELECT * FROM ' . self::table('vvd_clients') . ' WHERE id=%d', $id), ARRAY_A);
    }


    protected static function client_contacts($client) {
        $options = [];
        $seen = [];

        $push = function($email, $name = '', $meta = '') use (&$options, &$seen) {
            $email = sanitize_email((string) $email);
            if ($email === '') { return; }
            $key = strtolower($email);
            if (isset($seen[$key])) { return; }
            $label = trim((string) $name) !== '' ? trim((string) $name) : $email;
            if (trim((string) $meta) !== '') {
                $label .= ' · ' . trim((string) $meta);
            }
            $options[] = [
                'email' => $email,
                'name' => sanitize_text_field((string) $name),
                'label' => $label . ' <' . $email . '>',
            ];
            $seen[$key] = true;
        };

        $push($client['email'] ?? '', $client['contact_name'] ?? ($client['name'] ?? ''), 'Principal');

        $json = $client['contacts_json'] ?? '';
        $rows = json_decode((string) $json, true);
        if (is_array($rows)) {
            foreach ($rows as $row) {
                if (!is_array($row)) { continue; }
                $meta = trim(implode(' · ', array_filter([
                    $row['designation'] ?? '',
                    $row['department'] ?? '',
                    $row['contact_type'] ?? '',
                ])));
                $push($row['email'] ?? '', $row['name'] ?? '', $meta);
            }
        }

        return $options;
    }

    protected static function append_client_contact($existing_json, $contact) {
        $contacts = json_decode((string) $existing_json, true);
        if (!is_array($contacts)) {
            $contacts = [];
        }
        $email = sanitize_email((string) ($contact['email'] ?? ''));
        if ($email === '') {
            return wp_json_encode($contacts);
        }
        $normalized = [
            'id' => sanitize_text_field((string) ($contact['id'] ?? '')),
            'email' => $email,
            'name' => sanitize_text_field((string) ($contact['name'] ?? '')),
            'salutation' => sanitize_text_field((string) ($contact['salutation'] ?? '')),
            'first_name' => sanitize_text_field((string) ($contact['first_name'] ?? '')),
            'last_name' => sanitize_text_field((string) ($contact['last_name'] ?? '')),
            'phone' => sanitize_text_field((string) ($contact['phone'] ?? '')),
            'mobile' => sanitize_text_field((string) ($contact['mobile'] ?? '')),
            'designation' => sanitize_text_field((string) ($contact['designation'] ?? '')),
            'department' => sanitize_text_field((string) ($contact['department'] ?? '')),
            'contact_type' => sanitize_text_field((string) ($contact['contact_type'] ?? '')),
            'is_primary' => !empty($contact['is_primary']) ? 1 : 0,
        ];
        $replaced = false;
        foreach ($contacts as $index => $row) {
            if (!is_array($row)) { continue; }
            if (strtolower((string) ($row['email'] ?? '')) === strtolower($email)) {
                $contacts[$index] = array_merge($row, array_filter($normalized, function($v){ return $v !== '' && $v !== null; }));
                $replaced = true;
                break;
            }
        }
        if (!$replaced) {
            $contacts[] = $normalized;
        }
        return wp_json_encode(array_values($contacts));
    }

    protected static function invoice($id) {
        global $wpdb;
        $invoices = self::table('vvd_invoices');
        $clients = self::table('vvd_clients');
        return (array) $wpdb->get_row($wpdb->prepare(
            "SELECT i.*, c.name as client_name, c.tax_id as client_tax_id, c.address as client_address, c.email as client_email, c.phone as client_phone,
                    oi.series as original_series, oi.number as original_number, oi.code as original_code
             FROM {$invoices} i
             LEFT JOIN {$clients} c ON c.id = i.client_id
             LEFT JOIN {$invoices} oi ON oi.id = i.original_invoice_id
             WHERE i.id=%d",
            $id
        ), ARRAY_A);
    }

    protected static function invoice_lines($invoice_id) {
        global $wpdb;
        return $wpdb->get_results($wpdb->prepare('SELECT * FROM ' . self::table('vvd_invoice_lines') . ' WHERE invoice_id=%d ORDER BY id ASC', $invoice_id), ARRAY_A);
    }

    protected static function invoice_events($invoice_id) {
        global $wpdb;
        return $wpdb->get_results($wpdb->prepare('SELECT * FROM ' . self::table('vvd_events') . ' WHERE invoice_id=%d ORDER BY created_at DESC, id DESC', $invoice_id), ARRAY_A);
    }

    protected static function log_event($type, $message, $invoice_id = null, $client_id = null, $context = []) {
        global $wpdb;
        $wpdb->insert(self::table('vvd_events'), [
            'invoice_id' => $invoice_id ?: null,
            'client_id' => $client_id ?: null,
            'event_type' => sanitize_text_field($type),
            'event_message' => wp_strip_all_tags($message),
            'context' => !empty($context) ? wp_json_encode($context) : null,
            'user_id' => get_current_user_id(),
            'created_at' => current_time('mysql'),
        ]);
    }

    protected static function invoice_code($series, $number, $padding = 6) {
        return strtoupper($series) . '-' . str_pad((int) $number, (int) $padding, '0', STR_PAD_LEFT);
    }

    protected static function doc_label($type) {
        if ($type === 'rectificativa') return 'Rectificativa';
        if ($type === 'quote') return 'Presupuesto';
        return 'Factura';
    }

    protected static function public_document_token($invoice) {
        $code = is_array($invoice) ? (string) ($invoice['code'] ?? '') : '';
        $issue_date = is_array($invoice) ? (string) ($invoice['issue_date'] ?? '') : '';
        $id = is_array($invoice) ? (int) ($invoice['id'] ?? 0) : 0;
        return hash_hmac('sha256', $id . '|' . $code . '|' . $issue_date, wp_salt('auth'));
    }

    protected static function public_document_url($invoice_id, $invoice = null) {
        if (!$invoice) {
            $invoice = self::invoice((int) $invoice_id);
        }
        if (empty($invoice)) {
            return '';
        }
        $args = [
            'action' => 'vvd_public_document',
            'invoice_id' => (int) $invoice_id,
            'token' => self::public_document_token($invoice),
        ];
        return add_query_arg($args, admin_url('admin-post.php'));
    }

    protected static function verify_public_document_token($invoice, $token) {
        if (empty($invoice) || empty($token)) {
            return false;
        }
        $expected = self::public_document_token($invoice);
        return hash_equals($expected, (string) $token);
    }

    protected static function reserve_series_number($invoice_type) {
        global $wpdb;
        $settings = self::settings();
        if ($invoice_type === 'rectificativa') {
            $series_key = strtoupper($settings['rectificative_series'] ?: 'RT');
        } elseif ($invoice_type === 'quote') {
            $series_key = strtoupper($settings['quote_series'] ?: 'P');
        } else {
            $series_key = strtoupper($settings['default_series'] ?: 'A');
        }

        $table = self::table('vvd_series');
        $row = $wpdb->get_row($wpdb->prepare('SELECT * FROM ' . $table . ' WHERE series_key=%s LIMIT 1', $series_key), ARRAY_A);
        if (!$row) {
            $wpdb->insert($table, [
                'series_key' => $series_key,
                'label_text' => 'Serie ' . $series_key,
                'prefix' => $series_key,
                'padding' => (int) ($settings['series_padding'] ?: 6),
                'current_number' => 0,
                'invoice_type' => $invoice_type,
                'is_active' => 1,
                'created_at' => current_time('mysql'),
            ]);
            $row = $wpdb->get_row($wpdb->prepare('SELECT * FROM ' . $table . ' WHERE series_key=%s LIMIT 1', $series_key), ARRAY_A);
        }

        $next = ((int) $row['current_number']) + 1;
        $wpdb->update($table, [
            'current_number' => $next,
            'updated_at' => current_time('mysql'),
        ], ['id' => (int) $row['id']]);

        return [
            'series' => $series_key,
            'series_label' => $row['label_text'] ?: ('Serie ' . $series_key),
            'number' => $next,
            'padding' => (int) ($row['padding'] ?: $settings['series_padding'] ?: 6),
            'code' => self::invoice_code($series_key, $next, (int) ($row['padding'] ?: $settings['series_padding'] ?: 6)),
        ];
    }

    public static function admin_menu() {
        add_menu_page('Facturación VexelDot', 'Facturación', 'manage_options', 'vvd_facturas', [__CLASS__, 'page_invoices'], 'dashicons-media-spreadsheet', 26);
        add_submenu_page('vvd_facturas', 'Facturas y presupuestos', 'Facturación', 'manage_options', 'vvd_facturas', [__CLASS__, 'page_invoices']);
        add_submenu_page('vvd_facturas', 'Nueva factura', 'Nueva factura', 'manage_options', 'vvd_nueva_factura', [__CLASS__, 'page_invoice_form']);
        add_submenu_page('vvd_facturas', 'Presupuestos', 'Presupuestos', 'manage_options', 'vvd_presupuestos', [__CLASS__, 'page_quotes']);
        add_submenu_page('vvd_facturas', 'Nuevo presupuesto', 'Nuevo presupuesto', 'manage_options', 'vvd_nuevo_presupuesto', [__CLASS__, 'page_quote_form']);
        add_submenu_page('vvd_facturas', 'Clientes', 'Clientes', 'manage_options', 'vvd_clientes', [__CLASS__, 'page_clients']);
        add_submenu_page('vvd_facturas', 'Nuevo cliente', 'Nuevo cliente', 'manage_options', 'vvd_nuevo_cliente', [__CLASS__, 'page_client_form']);
        add_submenu_page('vvd_facturas', 'Ajustes', 'Ajustes', 'manage_options', 'vvd_ajustes', [__CLASS__, 'page_settings']);
        add_submenu_page('vvd_facturas', 'Análisis', 'Análisis', 'manage_options', 'vvd_analisis', [__CLASS__, 'page_analysis']);
        add_submenu_page('vvd_facturas', 'Libro de facturas', 'Libro de facturas', 'manage_options', 'vvd_libro_facturas', [__CLASS__, 'page_invoice_book']);
        add_submenu_page('vvd_facturas', 'Importar CSV', 'Importaciones', 'manage_options', 'vvd_importar_csv', [__CLASS__, 'page_import_accounting']);
        add_submenu_page('vvd_facturas', 'Importar clientes', 'Importar clientes', 'manage_options', 'vvd_importar_clientes', [__CLASS__, 'page_import_clients']);
        add_submenu_page(null, 'Enviar documento', 'Enviar documento', 'manage_options', 'vvd_enviar_documento', [__CLASS__, 'page_send_document']);
    }

    public static function enqueue($hook) {
        if (strpos($hook, 'vvd_') === false) {
            return;
        }
        wp_register_style('vvd-admin-inline', false);
        wp_enqueue_style('vvd-admin-inline');
        wp_add_inline_style('vvd-admin-inline', self::admin_css());
        wp_add_inline_script('jquery-core', self::admin_js());
    }

    protected static function admin_css() {
        return ':root{--vvd-primary:#6997c1;--vvd-secondary:#8d8e8e;--vvd-bg:#ffffff;--vvd-border:#d9e2ef;--vvd-soft:#f5f9fd}
        .vvd-wrap{max-width:1280px}.vvd-card{background:#fff;border:1px solid var(--vvd-border);border-radius:18px;padding:24px;margin:18px 0;box-shadow:0 10px 30px rgba(34,57,90,.06)}
        .vvd-grid{display:grid;grid-template-columns:repeat(12,1fr);gap:16px}.vvd-col-3{grid-column:span 3}.vvd-col-4{grid-column:span 4}.vvd-col-6{grid-column:span 6}.vvd-col-8{grid-column:span 8}.vvd-col-12{grid-column:span 12}
        .vvd-kpis{display:flex;gap:12px;flex-wrap:wrap;margin-bottom:14px}.vvd-kpi{display:inline-block;min-width:180px;padding:18px 20px;border-radius:18px;background:var(--vvd-primary);color:#fff}.vvd-kpi span{display:block;opacity:.8;font-size:12px;margin-bottom:3px}
        .vvd-table{width:100%;border-collapse:collapse}.vvd-table th,.vvd-table td{padding:12px 10px;border-bottom:1px solid #e7edf5;vertical-align:top}.vvd-table th{text-align:left;color:var(--vvd-secondary)}
        .vvd-btn{display:inline-block;padding:10px 14px;border-radius:12px;background:var(--vvd-primary);color:#fff;text-decoration:none;border:none;cursor:pointer}.vvd-btn.alt{background:var(--vvd-secondary)}.vvd-btn.light{background:#fff;color:var(--vvd-primary);border:1px solid var(--vvd-border)}.vvd-btn.danger{background:#b44343}
        .vvd-actions{display:flex;gap:8px;flex-wrap:wrap}.vvd-label{font-weight:600;color:#374151;display:block;margin-bottom:6px}.vvd-input,.vvd-textarea,.vvd-select{width:100%;padding:11px 12px;border:1px solid var(--vvd-border);border-radius:12px;box-sizing:border-box}.vvd-textarea{min-height:100px}
        .vvd-note{background:#f7fbff;border:1px solid var(--vvd-border);padding:12px;border-radius:12px}.vvd-lines .line{display:grid;grid-template-columns:minmax(0,1fr) 120px 160px 50px;gap:12px;align-items:start;margin-bottom:12px}.vvd-totals{margin-top:18px;padding:20px;border-radius:18px;background:var(--vvd-soft);border:1px solid var(--vvd-border);max-width:420px}
        .vvd-hero{background:linear-gradient(135deg,#6997c1 0%, #8d8e8e 100%);color:#fff;padding:28px;border-radius:20px;margin:16px 0}.vvd-hero h1,.vvd-hero h2{color:#fff;margin:0 0 10px 0}.vvd-code{font-family:Consolas,Monaco,monospace;background:#f6f8fa;padding:4px 6px;border-radius:6px}.vvd-meta{display:grid;grid-template-columns:repeat(4,1fr);gap:12px}.vvd-meta .item{background:#fff;border-radius:14px;padding:16px;border:1px solid rgba(255,255,255,.35);color:#27405d}.vvd-meta .item strong{display:block;color:#1d3047;margin-bottom:4px}.vvd-pill{display:inline-block;padding:4px 10px;border-radius:999px;background:#eef5fb;color:#315476;font-size:12px;font-weight:600}
        .vvd-detail-sections{display:grid;grid-template-columns:1.1fr .9fr;gap:18px}.vvd-panel{background:#fff;border:1px solid var(--vvd-border);border-radius:18px;padding:18px}.vvd-events{max-height:460px;overflow:auto}.vvd-event{padding:12px 0;border-bottom:1px solid #eef2f7}.vvd-event:last-child{border-bottom:0}
        .vvd-chart-grid{display:grid;grid-template-columns:2fr 1fr;gap:18px;margin:18px 0}.vvd-chart-card{background:#fff;border:1px solid var(--vvd-border);border-radius:20px;padding:22px;box-shadow:0 10px 30px rgba(34,57,90,.06)}.vvd-chart-header{display:flex;justify-content:space-between;align-items:flex-start;gap:12px;margin-bottom:10px}.vvd-chart-header h2{margin:0 0 4px}.vvd-chart-header p{margin:0;color:#6b7280}.vvd-chart-legend{display:flex;gap:16px;flex-wrap:wrap;margin:10px 0 12px}.vvd-chart-legend span{display:flex;align-items:center;gap:8px;color:#4b5563}.vvd-chart-legend i{display:inline-block;width:12px;height:12px;border-radius:4px}.vvd-chart-legend i.income{background:#12b981}.vvd-chart-legend i.expense{background:#ef4444}.vvd-chart-summary{display:flex;gap:18px;flex-wrap:wrap;margin-bottom:12px}.vvd-chart-summary div{padding:12px 14px;background:var(--vvd-soft);border:1px solid var(--vvd-border);border-radius:14px;min-width:180px}.vvd-chart-summary span{display:block;font-size:12px;color:#6b7280;margin-bottom:4px}.vvd-chart-summary strong{font-size:24px;line-height:1.1}.vvd-bars-wrap{display:flex;align-items:flex-end;gap:12px;min-height:220px;padding:14px 0 6px;border-bottom:1px dashed #d5deea;overflow-x:auto}.vvd-bar-group{min-width:52px;text-align:center}.vvd-bar-stack{height:190px;display:flex;align-items:flex-end;justify-content:center;gap:6px}.vvd-bar{width:18px;border-radius:10px 10px 4px 4px;min-height:2px}.vvd-bar.income{background:linear-gradient(180deg,#38d39f,#12b981)}.vvd-bar.expense{background:linear-gradient(180deg,#fb7185,#ef4444)}.vvd-bar-label{font-size:11px;color:#6b7280;margin-top:8px;text-transform:lowercase}.vvd-chart-foot{margin:12px 0 0;color:#6b7280}
        .vvd-balance-card{background:#fff;border:1px solid var(--vvd-border);border-radius:20px;padding:22px;box-shadow:0 10px 30px rgba(34,57,90,.06)}.vvd-balance-big{font-size:42px;line-height:1;color:#12b981;margin:6px 0 18px;font-weight:800}.vvd-balance-list{display:grid;gap:10px;margin-bottom:18px}.vvd-balance-item{display:flex;justify-content:space-between;gap:10px;padding:12px 0;border-bottom:1px dashed #d9e2ef}.vvd-balance-item:last-child{border-bottom:0}.vvd-balance-item span{color:#6b7280}.vvd-balance-item strong.income{color:#12b981}.vvd-balance-item strong.expense{color:#ef4444}
        .vvd-history-wrap{display:flex;align-items:flex-end;gap:14px;min-height:260px;overflow-x:auto;padding-top:12px}.vvd-history-col{min-width:68px;text-align:center}.vvd-history-bar{width:44px;margin:0 auto;border-radius:14px 14px 4px 4px;min-height:6px;background:linear-gradient(180deg,#6fd6b2,#12b981)}.vvd-history-bar.negative{background:linear-gradient(180deg,#f59e0b,#f97316)}.vvd-history-year{margin-top:10px;font-weight:700;color:#374151}.vvd-history-amount{font-size:11px;color:#6b7280;margin-top:4px}.vvd-filter-row{display:flex;gap:10px;align-items:end;flex-wrap:wrap;margin:12px 0 16px}.vvd-filter-row .vvd-field{min-width:180px;flex:1}.vvd-muted{color:#6b7280;font-size:12px}.vvd-type-sale{color:#12b981;font-weight:700}.vvd-type-purchase{color:#ef4444;font-weight:700}
        @media(max-width:960px){.vvd-col-3,.vvd-col-4,.vvd-col-6,.vvd-col-8,.vvd-meta .item{grid-column:span 12}.vvd-grid,.vvd-meta,.vvd-detail-sections,.vvd-chart-grid{grid-template-columns:1fr}.vvd-lines .line{grid-template-columns:1fr}.vvd-actions{flex-direction:column}.vvd-balance-big{font-size:34px}}';
    }

    protected static function admin_js() {
        return <<<'JS'
jQuery(function($){
    function recalc(){
        let base=0;
        $('.vvd-line').each(function(){
            const qty=parseFloat(String($(this).find('.vvd-qty').val()).replace(',', '.'))||0;
            const price=parseFloat(String($(this).find('.vvd-price').val()).replace(',', '.'))||0;
            const total=qty*price;
            $(this).find('.vvd-line-total').text(total.toFixed(2)+' €');
            base+=total;
        });
        const iva=parseFloat(String($('#tax_rate').val()).replace(',', '.'))||0;
        const irpf=parseFloat(String($('#irpf_rate').val()).replace(',', '.'))||0;
        const type=String($('#invoice_type').val()||'standard');
        const tax=base*iva/100;
        const irpfAmount=base*irpf/100;
        const total=type==='quote' ? base : base+tax-irpfAmount;
        $('#vvd-base').text(base.toFixed(2)+' €');
        $('#vvd-tax').text((type==='quote'?0:tax).toFixed(2)+' €');
        $('#vvd-irpf').text((type==='quote'?0:irpfAmount).toFixed(2)+' €');
        $('#vvd-total').text(total.toFixed(2)+' €');
    }
    $(document).on('input', '.vvd-qty,.vvd-price,#tax_rate,#irpf_rate', recalc);
    $(document).on('change', '#invoice_type', recalc);
    $(document).on('click', '.vvd-add-line', function(e){
        e.preventDefault();
        const row=$('.vvd-line-template').first().clone().removeClass('vvd-line-template').addClass('vvd-line').show();
        row.find('input').val('');
        row.find('.vvd-qty').val('1');
        $('.vvd-lines').append(row);
        recalc();
    });
    $(document).on('click', '.vvd-remove-line', function(e){
        e.preventDefault();
        $(this).closest('.line').remove();
        if(!$('.vvd-line').length){$('.vvd-add-line').trigger('click');}
        recalc();
    });
    recalc();
});
JS;
    }

    public static function page_invoices() {
        global $wpdb;
        if (isset($_GET['view'])) {
            self::page_invoice_detail((int) $_GET['view']);
            return;
        }

        $clients_rows = $wpdb->get_results('SELECT id, name FROM ' . self::table('vvd_clients') . ' ORDER BY name ASC', ARRAY_A);
        $filter_client_id = isset($_GET['client_id']) ? (int) $_GET['client_id'] : 0;
        $filter_paid = sanitize_text_field($_GET['paid_filter'] ?? 'all');
        if (!in_array($filter_paid, ['all', 'paid', 'unpaid'], true)) {
            $filter_paid = 'all';
        }

        $sql = "SELECT i.*, c.name as client_name FROM " . self::table('vvd_invoices') . " i LEFT JOIN " . self::table('vvd_clients') . " c ON c.id=i.client_id WHERE i.invoice_type != 'quote'";
        $params = [];
        if ($filter_client_id > 0) {
            $sql .= " AND i.client_id = %d";
            $params[] = $filter_client_id;
        }
        if ($filter_paid === 'paid') {
            $sql .= " AND (i.status = 'paid' OR i.paid_at IS NOT NULL)";
        } elseif ($filter_paid === 'unpaid') {
            $sql .= " AND (i.status != 'paid' AND i.paid_at IS NULL)";
        }
        $sql .= " ORDER BY i.issue_date DESC, i.id DESC";
        $rows = !empty($params) ? $wpdb->get_results($wpdb->prepare($sql, $params), ARRAY_A) : $wpdb->get_results($sql, ARRAY_A);

        $count = count($rows);
        $total = 0;
        $sent = 0;
        $paid = 0;
        foreach ($rows as $row) {
            $total += (float) $row['total_amount'];
            if (!empty($row['sent_at'])) { $sent++; }
            if (!empty($row['paid_at']) || (($row['status'] ?? '') === 'paid')) { $paid++; }
        }

        echo '<div class="wrap vvd-wrap">';
        echo '<div class="vvd-hero"><h1>Facturación VexelDot</h1><p>Solo facturas y rectificativas. Los presupuestos viven aparte para que no se mezcle lo comercial con lo contable.</p></div>';
        echo '<div class="vvd-kpis">';
        echo '<div class="vvd-kpi"><span>Total facturas</span><strong>' . esc_html($count) . '</strong></div>';
        echo '<div class="vvd-kpi"><span>Importe acumulado</span><strong>' . esc_html(number_format($total, 2, ',', '.')) . ' €</strong></div>';
        echo '<div class="vvd-kpi"><span>Enviadas</span><strong>' . esc_html($sent) . '</strong></div>';
        echo '<div class="vvd-kpi"><span>Pagadas</span><strong>' . esc_html($paid) . '</strong></div>';
        echo '</div>';
        echo '<p><a class="vvd-btn" href="' . esc_url(admin_url('admin.php?page=vvd_nueva_factura')) . '">+ Nueva factura</a> <a class="vvd-btn light" href="' . esc_url(admin_url('admin.php?page=vvd_presupuestos')) . '">Ver presupuestos</a> <a class="vvd-btn light" href="' . esc_url(admin_url('admin.php?page=vvd_nuevo_presupuesto')) . '">+ Nuevo presupuesto</a></p>';
        echo '<div class="vvd-card">';
        if (!empty($_GET['bulk_done'])) {
            echo '<div class="notice notice-success" style="margin:0 0 16px 0"><p>Actualización masiva aplicada correctamente.</p></div>';
        }
        echo '<form method="get" class="vvd-filter-row">';
        echo '<input type="hidden" name="page" value="vvd_facturas">';
        echo '<div class="vvd-field"><label class="vvd-label">Cliente</label><select name="client_id" class="vvd-select"><option value="0">Todos los clientes</option>';
        foreach ($clients_rows as $client_row) {
            echo '<option value="' . esc_attr((string) $client_row['id']) . '"' . selected($filter_client_id, (int) $client_row['id'], false) . '>' . esc_html($client_row['name']) . '</option>';
        }
        echo '</select></div>';
        echo '<div class="vvd-field"><label class="vvd-label">Pagadas</label><select name="paid_filter" class="vvd-select"><option value="all"' . selected($filter_paid, 'all', false) . '>Todas</option><option value="paid"' . selected($filter_paid, 'paid', false) . '>Solo pagadas</option><option value="unpaid"' . selected($filter_paid, 'unpaid', false) . '>Solo no pagadas</option></select></div>';
        echo '<div class="vvd-field" style="flex:0 0 auto"><label class="vvd-label">&nbsp;</label><button class="vvd-btn" type="submit">Filtrar</button></div>';
        echo '</form>';

        echo '<form method="post" action="' . esc_url(admin_url('admin-post.php')) . '" class="vvd-bulk-form">';
        wp_nonce_field('vvd_bulk_update_invoices');
        echo '<input type="hidden" name="action" value="vvd_bulk_update_invoices">';
        echo '<input type="hidden" name="client_id" value="' . esc_attr($filter_client_id) . '">';
        echo '<input type="hidden" name="paid_filter" value="' . esc_attr($filter_paid) . '">';
        echo '<div class="vvd-filter-row" style="margin:0 0 16px 0;align-items:flex-end">';
        echo '<div class="vvd-field"><label class="vvd-label">Acción masiva</label><select name="bulk_action" class="vvd-select"><option value="">Selecciona una acción</option><option value="mark_sent">Marcar enviadas</option><option value="mark_unsent">Marcar no enviadas</option><option value="mark_paid">Marcar pagadas</option><option value="mark_unpaid">Marcar no pagadas</option></select></div>';
        echo '<div class="vvd-field" style="flex:0 0 auto"><label class="vvd-label">&nbsp;</label><button class="vvd-btn alt" type="submit">Aplicar a seleccionadas</button></div>';
        echo '</div>';

        echo '<table class="vvd-table"><thead><tr><th style="width:34px"><input type="checkbox" class="vvd-check-all" aria-label="Seleccionar todas"></th><th>Documento</th><th>Fecha</th><th>Cliente</th><th>Tipo</th><th>Enviada</th><th>Pagada</th><th>VERI*FACTU</th><th>Total</th><th>Acciones</th></tr></thead><tbody>';
        foreach ($rows as $row) {
            $id = (int) $row['id'];
            $is_quote = ($row['invoice_type'] === 'quote');
            $code = $row['code'] ?: self::invoice_code($row['series'], $row['number'], (int) (self::settings()['series_padding'] ?: 6));
            $is_sent = !empty($row['sent_at']);
            $is_paid = !empty($row['paid_at']) || (($row['status'] ?? '') === 'paid');
            $toggle_label = $is_paid ? 'Marcar no pagada' : 'Marcar pagada';
            echo '<tr>';
            echo '<td><input type="checkbox" name="invoice_ids[]" value="' . esc_attr($id) . '"></td>';
            echo '<td><strong>' . esc_html($code) . '</strong><br><small>' . esc_html(self::invoice_status_label($row['status'])) . '</small></td>';
            echo '<td>' . esc_html($row['issue_date']) . '</td>';
            echo '<td>' . esc_html($row['client_name']) . '</td>';
            echo '<td>' . esc_html(self::doc_label($row['invoice_type'])) . '</td>';
            echo '<td>' . ($is_sent ? '<span class="vvd-pill">Sí</span><br><small class="vvd-muted">' . esc_html(date_i18n('d/m/Y', strtotime($row['sent_at']))) . '</small>' : '<span class="vvd-pill" style="background:#f3f4f6;color:#4b5563">No</span>') . '</td>';
            echo '<td>' . ($is_paid ? '<span class="vvd-pill" style="background:#dcfce7;color:#166534">Sí</span><br><small class="vvd-muted">' . esc_html(!empty($row['paid_at']) ? date_i18n('d/m/Y', strtotime($row['paid_at'])) : 'Importada') . '</small>' : '<span class="vvd-pill" style="background:#fef2f2;color:#991b1b">No</span>') . '</td>';
            echo '<td><span class="vvd-pill">' . esc_html($is_quote ? 'n/a' : $row['verifactu_status']) . '</span></td>';
            echo '<td><strong>' . esc_html(number_format((float) $row['total_amount'], 2, ',', '.')) . ' €</strong></td>';
            echo '<td><div class="vvd-actions">';
            echo '<a class="vvd-btn light" href="' . esc_url(admin_url('admin.php?page=vvd_facturas&view=' . $id)) . '">Ver</a>';
            echo '<a class="vvd-btn light" href="' . esc_url(admin_url('admin.php?page=vvd_nueva_factura&edit_id=' . $id)) . '">Editar</a>';
            echo '<a class="vvd-btn light" href="' . esc_url(admin_url('admin.php?page=vvd_nueva_factura&duplicate_id=' . $id)) . '">Duplicar</a>';
            if ($is_quote) {
                $convert_url = wp_nonce_url(admin_url('admin-post.php?action=vvd_convert_to_invoice&invoice_id=' . $id), 'vvd_convert_' . $id);
                echo '<a class="vvd-btn alt" href="' . esc_url($convert_url) . '">Convertir en factura</a>';
            } else {
                echo '<a class="vvd-btn light" href="' . esc_url(admin_url('admin.php?page=vvd_nueva_factura&rectify_id=' . $id)) . '">Rectificativa</a>';
                $toggle_url = wp_nonce_url(add_query_arg([
                    'action' => 'vvd_toggle_invoice_paid',
                    'invoice_id' => $id,
                    'client_id' => $filter_client_id,
                    'paid_filter' => $filter_paid,
                ], admin_url('admin-post.php')), 'vvd_toggle_paid_' . $id);
                echo '<a class="vvd-btn light" href="' . esc_url($toggle_url) . '">' . esc_html($toggle_label) . '</a>';
            }
            echo '<a class="vvd-btn alt" target="_blank" href="' . esc_url(wp_nonce_url(admin_url('admin-post.php?action=vvd_print_document&invoice_id=' . $id), 'vvd_print_' . $id)) . '">PDF</a>';
            $email_url = admin_url('admin.php?page=vvd_enviar_documento&invoice_id=' . $id);
            echo '<a class="vvd-btn alt" href="' . esc_url($email_url) . '">Email</a>';
            $delete_url = wp_nonce_url(admin_url('admin-post.php?action=vvd_delete_invoice&invoice_id=' . $id), 'vvd_delete_invoice_' . $id);
            echo '<a class="vvd-btn danger" onclick="return confirm(&quot;¿Eliminar documento?&quot;)" href="' . esc_url($delete_url) . '">Eliminar</a>';
            echo '</div></td>';
            echo '</tr>';
        }
        if (!$rows) {
            echo '<tr><td colspan="10">No hay documentos para el filtro seleccionado.</td></tr>';
        }
        echo '</tbody></table></form></div></div>';
        echo '<script>document.addEventListener("DOMContentLoaded",function(){var a=document.querySelector(".vvd-check-all");if(!a)return;a.addEventListener("change",function(){document.querySelectorAll("input[name=\"invoice_ids[]\"]").forEach(function(el){el.checked=a.checked;});});});</script>';
    }

    public static function bulk_update_invoices() {
        if (!current_user_can('manage_options')) {
            wp_die('No autorizado');
        }

        check_admin_referer('vvd_bulk_update_invoices');

        $invoice_ids = isset($_POST['invoice_ids']) ? array_map('intval', (array) $_POST['invoice_ids']) : [];
        $bulk_action = sanitize_text_field($_POST['bulk_action'] ?? '');
        $client_id = isset($_POST['client_id']) ? (int) $_POST['client_id'] : 0;
        $paid_filter = sanitize_text_field($_POST['paid_filter'] ?? 'all');

        if (!$invoice_ids || !in_array($bulk_action, ['mark_sent', 'mark_unsent', 'mark_paid', 'mark_unpaid'], true)) {
            wp_safe_redirect(add_query_arg([
                'page' => 'vvd_facturas',
                'client_id' => $client_id,
                'paid_filter' => $paid_filter,
            ], admin_url('admin.php')));
            exit;
        }

        global $wpdb;
        foreach ($invoice_ids as $invoice_id) {
            $invoice = $wpdb->get_row($wpdb->prepare('SELECT * FROM ' . self::table('vvd_invoices') . ' WHERE id=%d', $invoice_id), ARRAY_A);
            if (!$invoice || (($invoice['invoice_type'] ?? '') === 'quote')) {
                continue;
            }

            switch ($bulk_action) {
                case 'mark_sent':
                    $new_status = ((!empty($invoice['paid_at']) || (($invoice['status'] ?? '') === 'paid')) ? 'paid' : 'sent');
                    $wpdb->update(self::table('vvd_invoices'), [
                        'sent_at' => !empty($invoice['sent_at']) ? $invoice['sent_at'] : current_time('mysql'),
                        'status' => $new_status,
                    ], ['id' => $invoice_id]);
                    self::log_event('invoice_bulk_sent', 'Factura marcada como enviada masivamente.', $invoice_id, $invoice['client_id']);
                    break;
                case 'mark_unsent':
                    $new_status = (!empty($invoice['paid_at']) || (($invoice['status'] ?? '') === 'paid')) ? 'paid' : 'issued';
                    $wpdb->update(self::table('vvd_invoices'), [
                        'sent_at' => null,
                        'status' => $new_status,
                    ], ['id' => $invoice_id]);
                    self::log_event('invoice_bulk_unsent', 'Factura marcada como no enviada masivamente.', $invoice_id, $invoice['client_id']);
                    break;
                case 'mark_paid':
                    $wpdb->update(self::table('vvd_invoices'), [
                        'paid_at' => !empty($invoice['paid_at']) ? $invoice['paid_at'] : current_time('mysql'),
                        'status' => 'paid',
                        'sent_at' => !empty($invoice['sent_at']) ? $invoice['sent_at'] : current_time('mysql'),
                    ], ['id' => $invoice_id]);
                    self::log_event('invoice_bulk_paid', 'Factura marcada como pagada masivamente.', $invoice_id, $invoice['client_id']);
                    break;
                case 'mark_unpaid':
                    $new_status = !empty($invoice['sent_at']) ? 'sent' : 'issued';
                    $wpdb->update(self::table('vvd_invoices'), [
                        'paid_at' => null,
                        'status' => $new_status,
                    ], ['id' => $invoice_id]);
                    self::log_event('invoice_bulk_unpaid', 'Factura marcada como no pagada masivamente.', $invoice_id, $invoice['client_id']);
                    break;
            }
        }

        wp_safe_redirect(add_query_arg([
            'page' => 'vvd_facturas',
            'client_id' => $client_id,
            'paid_filter' => $paid_filter,
            'bulk_done' => 1,
        ], admin_url('admin.php')));
        exit;
    }

    public static function page_quotes() {
        global $wpdb;
        if (isset($_GET['view'])) {
            self::page_invoice_detail((int) $_GET['view']);
            return;
        }

        $rows = $wpdb->get_results("SELECT i.*, c.name as client_name FROM " . self::table('vvd_invoices') . " i LEFT JOIN " . self::table('vvd_clients') . " c ON c.id=i.client_id WHERE i.invoice_type = 'quote' ORDER BY i.issue_date DESC, i.id DESC", ARRAY_A);
        $count = count($rows);
        $total = 0;
        $sent = 0;
        foreach ($rows as $row) {
            $total += (float) $row['total_amount'];
            if (!empty($row['sent_at'])) { $sent++; }
        }

        echo '<div class="wrap vvd-wrap">';
        echo '<div class="vvd-hero"><h1>Presupuestos</h1><p>Zona separada para preparar, enviar y convertir presupuestos en factura sin mezclarlo con la facturación emitida.</p></div>';
        echo '<div class="vvd-kpis">';
        echo '<div class="vvd-kpi"><span>Total presupuestos</span><strong>' . esc_html($count) . '</strong></div>';
        echo '<div class="vvd-kpi"><span>Importe acumulado</span><strong>' . esc_html(number_format($total, 2, ',', '.')) . ' €</strong></div>';
        echo '<div class="vvd-kpi"><span>Enviados</span><strong>' . esc_html($sent) . '</strong></div>';
        echo '</div>';
        echo '<p><a class="vvd-btn" href="' . esc_url(admin_url('admin.php?page=vvd_nuevo_presupuesto')) . '">+ Nuevo presupuesto</a> <a class="vvd-btn light" href="' . esc_url(admin_url('admin.php?page=vvd_facturas')) . '">Ver facturación</a></p>';
        echo '<div class="vvd-card"><table class="vvd-table"><thead><tr><th>Presupuesto</th><th>Fecha</th><th>Cliente</th><th>Estado</th><th>Total</th><th>Acciones</th></tr></thead><tbody>';
        foreach ($rows as $row) {
            $id = (int) $row['id'];
            $code = $row['code'] ?: self::invoice_code($row['series'], $row['number'], (int) (self::settings()['series_padding'] ?: 6));
            echo '<tr>';
            echo '<td><strong>' . esc_html($code) . '</strong></td>';
            echo '<td>' . esc_html($row['issue_date']) . '</td>';
            echo '<td>' . esc_html($row['client_name']) . '</td>';
            echo '<td>' . esc_html(self::invoice_status_label($row['status'])) . '</td>';
            echo '<td><strong>' . esc_html(number_format((float) $row['total_amount'], 2, ',', '.')) . ' €</strong></td>';
            echo '<td><div class="vvd-actions">';
            echo '<a class="vvd-btn light" href="' . esc_url(admin_url('admin.php?page=vvd_presupuestos&view=' . $id)) . '">Ver</a>';
            echo '<a class="vvd-btn light" href="' . esc_url(admin_url('admin.php?page=vvd_nuevo_presupuesto&edit_id=' . $id)) . '">Editar</a>';
            echo '<a class="vvd-btn light" href="' . esc_url(admin_url('admin.php?page=vvd_nuevo_presupuesto&duplicate_id=' . $id)) . '">Duplicar</a>';
            echo '<form method="post" action="' . esc_url(admin_url('admin-post.php')) . '" style="display:inline">';
            wp_nonce_field('vvd_convert_' . $id);
            echo '<input type="hidden" name="action" value="vvd_convert_to_invoice"><input type="hidden" name="invoice_id" value="' . esc_attr($id) . '"><button class="vvd-btn alt" type="submit">Convertir en factura</button></form>';
            echo '<a class="vvd-btn alt" target="_blank" href="' . esc_url(wp_nonce_url(admin_url('admin-post.php?action=vvd_print_document&invoice_id=' . $id), 'vvd_print_' . $id)) . '">PDF</a>';
            echo '<a class="vvd-btn alt" href="' . esc_url(admin_url('admin.php?page=vvd_enviar_documento&invoice_id=' . $id)) . '">Email</a>';
            echo '<form method="post" action="' . esc_url(admin_url('admin-post.php')) . '" style="display:inline" onsubmit="return confirm(&quot;¿Eliminar presupuesto?&quot;)">';
            wp_nonce_field('vvd_delete_invoice_' . $id);
            echo '<input type="hidden" name="action" value="vvd_delete_invoice"><input type="hidden" name="invoice_id" value="' . esc_attr($id) . '"><button class="vvd-btn danger" type="submit">Eliminar</button></form>';
            echo '</div></td>';
            echo '</tr>';
        }
        if (!$rows) {
            echo '<tr><td colspan="6">No hay presupuestos aún.</td></tr>';
        }
        echo '</tbody></table></div></div>';
    }

    public static function page_invoice_detail($invoice_id) {
        $invoice = self::invoice($invoice_id);
        if (!$invoice) {
            wp_die('Documento no encontrado');
        }
        $client = self::client($invoice['client_id']);
        $events = self::invoice_events($invoice_id);
        $lines = self::invoice_lines($invoice_id);
        $code = $invoice['code'] ?: self::invoice_code($invoice['series'], $invoice['number'], (int) (self::settings()['series_padding'] ?: 6));

        echo '<div class="wrap vvd-wrap">';
        echo '<div class="vvd-hero"><h1>' . esc_html(self::doc_label($invoice['invoice_type'])) . ' ' . esc_html($code) . '</h1><div class="vvd-meta">';
        echo '<div class="item"><strong>Cliente</strong>' . esc_html($client['name'] ?: '-') . '</div>';
        echo '<div class="item"><strong>Fecha</strong>' . esc_html($invoice['issue_date']) . '</div>';
        echo '<div class="item"><strong>Estado</strong>' . esc_html(self::invoice_status_label($invoice['status'])) . '</div>';
        echo '<div class="item"><strong>Total</strong>' . esc_html(number_format((float) $invoice['total_amount'], 2, ',', '.')) . ' €</div>';
        echo '</div></div>';

        $back_page = (($invoice['invoice_type'] ?? '') === 'quote') ? 'vvd_presupuestos' : 'vvd_facturas';
        $edit_page = (($invoice['invoice_type'] ?? '') === 'quote') ? 'vvd_nuevo_presupuesto' : 'vvd_nueva_factura';
        echo '<p><a class="vvd-btn light" href="' . esc_url(admin_url('admin.php?page=' . $back_page)) . '">← Volver</a> <a class="vvd-btn" href="' . esc_url(admin_url('admin.php?page=' . $edit_page . '&edit_id=' . $invoice_id)) . '">Editar</a> <a class="vvd-btn alt" href="' . esc_url(wp_nonce_url(admin_url('admin-post.php?action=vvd_print_document&invoice_id=' . $invoice_id), 'vvd_print_' . $invoice_id)) . '">PDF</a></p>';
        echo '<div class="vvd-detail-sections">';
        echo '<div class="vvd-panel"><h2>Resumen</h2><table class="vvd-table"><tbody>';
        echo '<tr><th>Código</th><td>' . esc_html($code) . '</td></tr>';
        echo '<tr><th>Fecha</th><td>' . esc_html($invoice['issue_date']) . '</td></tr>';
        echo '<tr><th>Tipo</th><td>' . esc_html(self::doc_label($invoice['invoice_type'])) . '</td></tr>';
        echo '<tr><th>Serie</th><td>' . esc_html($invoice['series_label'] ?: $invoice['series']) . '</td></tr>';
        echo '<tr><th>Hash actual</th><td><span class="vvd-code">' . esc_html($invoice['hash_current'] ?: '-') . '</span></td></tr>';
        echo '<tr><th>Hash previo</th><td><span class="vvd-code">' . esc_html($invoice['hash_prev'] ?: '-') . '</span></td></tr>';
        echo '<tr><th>ID externo</th><td>' . esc_html($invoice['verifactu_external_id'] ?: '-') . '</td></tr>';
        echo '<tr><th>Notas</th><td>' . nl2br(esc_html($invoice['notes'] ?: '-')) . '</td></tr>';
        echo '</tbody></table>';
        echo '<h3>Líneas</h3><table class="vvd-table"><thead><tr><th>Concepto</th><th>Cant.</th><th>Unitario</th><th>Total</th></tr></thead><tbody>';
        foreach ($lines as $line) {
            echo '<tr><td>' . esc_html($line['description']) . '</td><td>' . esc_html(number_format((float) $line['quantity'], 2, ',', '.')) . '</td><td>' . esc_html(number_format((float) $line['unit_price'], 2, ',', '.')) . ' €</td><td><strong>' . esc_html(number_format((float) $line['line_total'], 2, ',', '.')) . ' €</strong></td></tr>';
        }
        echo '</tbody></table></div>';

        echo '<div class="vvd-panel"><h2>Cliente y trazabilidad</h2><table class="vvd-table"><tbody>';
        echo '<tr><th>Cliente</th><td>' . esc_html($client['name'] ?: '-') . '</td></tr>';
        echo '<tr><th>NIF</th><td>' . esc_html($client['tax_id'] ?: '-') . '</td></tr>';
        echo '<tr><th>Email</th><td>' . esc_html($client['email'] ?: '-') . '</td></tr>';
        echo '<tr><th>Teléfono</th><td>' . esc_html($client['phone'] ?: '-') . '</td></tr>';
        echo '<tr><th>Base imponible</th><td>' . esc_html(number_format((float) $invoice['taxable_base'], 2, ',', '.')) . ' €</td></tr>';
        echo '<tr><th>IVA</th><td>' . esc_html(number_format((float) $invoice['tax_amount'], 2, ',', '.')) . ' €</td></tr>';
        echo '<tr><th>IRPF</th><td>' . esc_html(number_format((float) $invoice['irpf_amount'], 2, ',', '.')) . ' €</td></tr>';
        echo '<tr><th>Total</th><td><strong>' . esc_html(number_format((float) $invoice['total_amount'], 2, ',', '.')) . ' €</strong></td></tr>';
        echo '</tbody></table>';
        echo '<div class="vvd-note" style="margin-top:14px"><strong>Base VERI*FACTU real preparada:</strong><br>La estructura ya guarda estado, hash actual, hash previo e ID externo para poder conectar después el envío reglado y el encadenado real.</div>';
        echo '<h3 style="margin-top:20px">Histórico de eventos</h3><div class="vvd-events">';
        if ($events) {
            foreach ($events as $event) {
                echo '<div class="vvd-event"><strong>' . esc_html($event['event_type']) . '</strong><br><small>' . esc_html($event['created_at']) . '</small><p style="margin:8px 0 0 0">' . esc_html($event['event_message']) . '</p></div>';
            }
        } else {
            echo '<p>No hay eventos todavía.</p>';
        }
        echo '</div></div>';
        echo '</div></div>';
    }

    public static function page_clients() {
        global $wpdb;
        if (isset($_GET['edit_id'])) {
            self::page_client_form();
            return;
        }
        $rows = $wpdb->get_results('SELECT * FROM ' . self::table('vvd_clients') . ' ORDER BY name ASC', ARRAY_A);
        echo '<div class="wrap vvd-wrap"><h1>Clientes</h1><p><a class="vvd-btn" href="' . esc_url(admin_url('admin.php?page=vvd_nuevo_cliente')) . '">+ Nuevo cliente</a> <a class="vvd-btn light" href="' . esc_url(admin_url('admin.php?page=vvd_importar_clientes')) . '">Importar clientes</a></p>';
        echo '<div class="vvd-card"><table class="vvd-table"><thead><tr><th>Cliente</th><th>Contacto</th><th>Email principal</th><th>Teléfonos</th><th>Pago / Estado</th><th>Acciones</th></tr></thead><tbody>';
        foreach ($rows as $row) {
            $id = (int) $row['id'];
            $contact_label = trim(($row['first_name'] ?? '') . ' ' . ($row['last_name'] ?? ''));
            if ($contact_label === '') { $contact_label = $row['contact_name'] ?? ''; }
            $phones = trim(implode(' · ', array_filter([$row['phone'] ?? '', $row['mobile'] ?? ''])));
            echo '<tr><td><strong>' . esc_html($row['name']) . '</strong><br><small class="vvd-muted">' . esc_html($row['tax_id']) . '</small></td><td>' . esc_html($contact_label ?: '—') . '<br><small class="vvd-muted">' . esc_html($row['company_name'] ?: ($row['display_name'] ?: '')) . '</small></td><td>' . esc_html($row['email']) . '</td><td>' . esc_html($phones ?: '—') . '</td><td>' . esc_html($row['payment_terms_label'] ?: '—') . '<br><small class="vvd-muted">' . esc_html($row['client_status'] ?: '—') . '</small></td><td><div class="vvd-actions">';
            echo '<a class="vvd-btn light" href="' . esc_url(admin_url('admin.php?page=vvd_clientes&edit_id=' . $id)) . '">Editar</a>';
            echo '<form method="post" action="' . esc_url(admin_url('admin-post.php')) . '" style="display:inline" onsubmit="return confirm(&quot;¿Eliminar cliente?&quot;)">';
            wp_nonce_field('vvd_delete_client_' . $id);
            echo '<input type="hidden" name="action" value="vvd_delete_client"><input type="hidden" name="client_id" value="' . esc_attr($id) . '"><button class="vvd-btn danger" type="submit">Eliminar</button></form>';
            echo '</div></td></tr>';
        }
        if (!$rows) {
            echo '<tr><td colspan="6">No hay clientes aún.</td></tr>';
        }
        echo '</tbody></table></div></div>';
    }

    public static function page_client_form() {
        $client_id = isset($_GET['edit_id']) ? (int) $_GET['edit_id'] : 0;
        $client = $client_id ? self::client($client_id) : [];
        echo '<div class="wrap vvd-wrap"><h1>' . ($client_id ? 'Editar cliente' : 'Nuevo cliente') . '</h1><div class="vvd-card"><form method="post" action="' . esc_url(admin_url('admin-post.php')) . '">';
        wp_nonce_field('vvd_save_client');
        echo '<input type="hidden" name="action" value="vvd_save_client"><input type="hidden" name="client_id" value="' . esc_attr($client_id) . '"><div class="vvd-grid">';
        self::field('Nombre visible', 'name', $client['name'] ?? '');
        self::field('Display name', 'display_name', $client['display_name'] ?? '');
        self::field('Empresa', 'company_name', $client['company_name'] ?? '');
        self::field('NIF / CIF', 'tax_id', $client['tax_id'] ?? '');
        self::field('Contacto', 'contact_name', $client['contact_name'] ?? '');
        self::field('Tratamiento', 'salutation', $client['salutation'] ?? '');
        self::field('Nombre contacto', 'first_name', $client['first_name'] ?? '');
        self::field('Apellidos contacto', 'last_name', $client['last_name'] ?? '');
        self::field('Email principal', 'email', $client['email'] ?? '', 'email');
        self::field('Teléfono', 'phone', $client['phone'] ?? '');
        self::field('Móvil', 'mobile', $client['mobile'] ?? '');
        self::field('Web', 'website', $client['website'] ?? '', 'url');
        self::field('Estado', 'client_status', $client['client_status'] ?? '');
        self::field('Moneda', 'currency_code', $client['currency_code'] ?? 'EUR');
        self::field('Condición de pago', 'payment_terms_label', $client['payment_terms_label'] ?? '');
        self::field('Código condición pago', 'payment_terms', $client['payment_terms'] ?? '');
        self::field('Atención facturación', 'billing_attention', $client['billing_attention'] ?? '');
        self::field('Dirección facturación', 'address', $client['address'] ?? '', 'textarea', 'vvd-col-12');
        self::field('Billing Street2', 'billing_street2', $client['billing_street2'] ?? '', 'text', 'vvd-col-12');
        self::field('Ciudad', 'city', $client['city'] ?? '');
        self::field('Provincia', 'state', $client['state'] ?? '');
        self::field('CP', 'postcode', $client['postcode'] ?? '');
        self::field('País', 'country', $client['country'] ?? '');
        self::field('Atención envío', 'shipping_attention', $client['shipping_attention'] ?? '');
        self::field('Dirección envío', 'shipping_address', $client['shipping_address'] ?? '', 'textarea', 'vvd-col-12');
        self::field('Shipping Street2', 'shipping_street2', $client['shipping_street2'] ?? '', 'text', 'vvd-col-12');
        self::field('Ciudad envío', 'shipping_city', $client['shipping_city'] ?? '');
        self::field('Provincia envío', 'shipping_state', $client['shipping_state'] ?? '');
        self::field('CP envío', 'shipping_postcode', $client['shipping_postcode'] ?? '');
        self::field('País envío', 'shipping_country', $client['shipping_country'] ?? '');
        self::field('Department', 'department', $client['department'] ?? '');
        self::field('Cargo', 'designation', $client['designation'] ?? '');
        self::field('Tipo contacto', 'contact_type', $client['contact_type'] ?? '');
        self::field('Taxable', 'taxable', $client['taxable'] ?? '');
        self::field('Nombre impuesto', 'tax_name', $client['tax_name'] ?? '');
        self::field('Porcentaje impuesto', 'tax_percentage', $client['tax_percentage'] ?? '');
        self::field('Motivo exención', 'exemption_reason', $client['exemption_reason'] ?? '', 'textarea', 'vvd-col-12');
        self::field('Notas', 'notes', $client['notes'] ?? '', 'textarea', 'vvd-col-12');
        echo '</div>';
        $contacts = self::client_contacts($client);
        if ($contacts) {
            echo '<div class="vvd-card" style="margin-top:18px"><h2>Contactos disponibles para envío</h2><table class="vvd-table"><thead><tr><th>Nombre</th><th>Email</th></tr></thead><tbody>';
            foreach ($contacts as $contact) {
                echo '<tr><td>' . esc_html($contact['name'] ?: 'Contacto') . '</td><td>' . esc_html($contact['email']) . '</td></tr>';
            }
            echo '</tbody></table></div>';
        }
        echo '<p><button class="vvd-btn" type="submit">Guardar cliente</button></p></form></div></div>';
    }

    public static function page_settings() {
        $s = self::settings();
        $series_rows = self::series_rows();
        echo '<div class="wrap vvd-wrap"><h1>Ajustes</h1><div class="vvd-card"><form method="post" action="' . esc_url(admin_url('admin-post.php')) . '">';
        wp_nonce_field('vvd_save_settings');
        echo '<input type="hidden" name="action" value="vvd_save_settings"><div class="vvd-grid">';
        self::field('Nombre emisor', 'issuer_name', $s['issuer_name'] ?? '');
        self::field('NIF emisor', 'issuer_tax_id', $s['issuer_tax_id'] ?? '');
        self::field('Email emisor', 'issuer_email', $s['issuer_email'] ?? '', 'email');
        self::field('Teléfono emisor', 'issuer_phone', $s['issuer_phone'] ?? '');
        self::field('IBAN', 'issuer_iban', $s['issuer_iban'] ?? '');
        self::field('Serie estándar', 'default_series', $s['default_series'] ?? 'A');
        self::field('Serie rectificativas', 'rectificative_series', $s['rectificative_series'] ?? 'RT');
        self::field('Serie presupuestos', 'quote_series', $s['quote_series'] ?? 'P');
        self::field('Relleno numeración', 'series_padding', $s['series_padding'] ?? 6, 'number');
        self::field('IVA por defecto', 'default_tax_rate', $s['default_tax_rate'] ?? '21');
        self::field('IRPF por defecto', 'default_irpf_rate', $s['default_irpf_rate'] ?? '15');
        self::field('Logo URL', 'logo_url', $s['logo_url'] ?? '', 'url', 'vvd-col-12');
        self::field('Fondo factura URL', 'background_image_url', $s['background_image_url'] ?? '', 'url', 'vvd-col-12');
        self::field('Color principal', 'brand_primary', $s['brand_primary'] ?? '#6997c1');
        self::field('Color fondo', 'brand_background', $s['brand_background'] ?? '#FFFFFF');
        self::field('Color secundario', 'brand_secondary', $s['brand_secondary'] ?? '#8d8e8e');
        self::field('Dirección emisor', 'issuer_address', $s['issuer_address'] ?? '', 'textarea', 'vvd-col-12');
        echo '</div><p><button class="vvd-btn" type="submit">Guardar ajustes</button></p></form></div>';

        echo '<div class="vvd-card"><h2>Series y numeración</h2><table class="vvd-table"><thead><tr><th>Serie</th><th>Etiqueta</th><th>Tipo</th><th>Actual</th><th>Padding</th></tr></thead><tbody>';
        foreach ($series_rows as $row) {
            echo '<tr><td><strong>' . esc_html($row['series_key']) . '</strong></td><td>' . esc_html($row['label_text']) . '</td><td>' . esc_html($row['invoice_type']) . '</td><td>' . esc_html($row['current_number']) . '</td><td>' . esc_html($row['padding']) . '</td></tr>';
        }
        echo '</tbody></table></div></div>';
    }

    public static function page_quote_form() {
        $_GET['force_type'] = 'quote';
        self::page_invoice_form();
    }

    public static function page_invoice_form() {
        global $wpdb;
        $clients = $wpdb->get_results('SELECT * FROM ' . self::table('vvd_clients') . ' ORDER BY name ASC', ARRAY_A);
        $settings = self::settings();
        $edit_id = isset($_GET['edit_id']) ? (int) $_GET['edit_id'] : 0;
        $duplicate_id = isset($_GET['duplicate_id']) ? (int) $_GET['duplicate_id'] : 0;
        $rectify_id = isset($_GET['rectify_id']) ? (int) $_GET['rectify_id'] : 0;
        $forced_type = sanitize_text_field($_GET['force_type'] ?? '');

        $title = 'Nueva factura';
        $invoice_type = $forced_type ?: 'standard';
        $selected_client_id = 0;
        $issue_date = current_time('Y-m-d');
        $tax_rate = $settings['default_tax_rate'] ?? 21;
        $irpf_rate = $settings['default_irpf_rate'] ?? 15;
        $notes = '';
        $original_invoice_id = 0;
        $lines = [['description' => '', 'quantity' => '1', 'unit_price' => '']];

        if ($edit_id) {
            $invoice = self::invoice($edit_id);
            if ($invoice) {
                $title = 'Editar ' . strtolower(self::doc_label($invoice['invoice_type']));
                $invoice_type = $invoice['invoice_type'];
                $selected_client_id = (int) $invoice['client_id'];
                $issue_date = $invoice['issue_date'];
                $tax_rate = $invoice['tax_rate'];
                $irpf_rate = $invoice['irpf_rate'];
                $notes = $invoice['notes'];
                $original_invoice_id = (int) $invoice['original_invoice_id'];
                $lines = self::invoice_lines($edit_id);
            }
        } elseif ($duplicate_id || $rectify_id) {
            $source_invoice = self::invoice($duplicate_id ?: $rectify_id);
            if ($source_invoice) {
                $selected_client_id = (int) $source_invoice['client_id'];
                $issue_date = current_time('Y-m-d');
                $notes = $source_invoice['notes'];
                $tax_rate = $source_invoice['tax_rate'];
                $irpf_rate = $source_invoice['irpf_rate'];
                $lines = self::invoice_lines($source_invoice['id']);
                if ($rectify_id) {
                    $invoice_type = 'rectificativa';
                    $original_invoice_id = (int) $source_invoice['id'];
                    foreach ($lines as &$line) {
                        $line['unit_price'] = (float) $line['unit_price'] * -1;
                    }
                    unset($line);
                    $notes = trim("Rectificativa del documento " . ($source_invoice['code'] ?: self::invoice_code($source_invoice['series'], $source_invoice['number'], (int) ($settings['series_padding'] ?: 6))) . "
" . $notes);
                    $title = 'Nueva rectificativa';
                } else {
                    $invoice_type = $source_invoice['invoice_type'];
                    $title = 'Duplicar ' . strtolower(self::doc_label($source_invoice['invoice_type']));
                }
            }
        } elseif ($forced_type === 'quote') {
            $title = 'Nuevo presupuesto';
        }

        echo '<div class="wrap vvd-wrap"><h1>' . esc_html($title) . '</h1>';
        if (!$clients) {
            echo '<div class="vvd-card"><p class="vvd-note">Primero necesitas crear un cliente para poder emitir documentos.</p><p><a class="vvd-btn" href="' . esc_url(admin_url('admin.php?page=vvd_nuevo_cliente')) . '">Crear cliente</a></p></div></div>';
            return;
        }

        echo '<div class="vvd-card"><form method="post" action="' . esc_url(admin_url('admin-post.php')) . '">';
        wp_nonce_field('vvd_save_invoice');
        echo '<input type="hidden" name="action" value="vvd_save_invoice"><input type="hidden" name="invoice_id" value="' . esc_attr($edit_id) . '"><div class="vvd-grid">';
        echo '<div class="vvd-col-4"><label class="vvd-label">Cliente</label><select class="vvd-select" name="client_id">';
        foreach ($clients as $client) {
            echo '<option value="' . esc_attr($client['id']) . '" ' . selected($selected_client_id, $client['id'], false) . '>' . esc_html($client['name']) . '</option>';
        }
        echo '</select></div>';
        self::field('Fecha', 'issue_date', $issue_date, 'date', 'vvd-col-4');
        echo '<div class="vvd-col-4"><label class="vvd-label">Tipo</label><select class="vvd-select" id="invoice_type" name="invoice_type"><option value="standard" ' . selected($invoice_type, 'standard', false) . '>Factura</option><option value="rectificativa" ' . selected($invoice_type, 'rectificativa', false) . '>Rectificativa</option><option value="quote" ' . selected($invoice_type, 'quote', false) . '>Presupuesto</option></select></div>';
        self::field('Documento original ID', 'original_invoice_id', $original_invoice_id, 'number');
        self::field('IVA (%)', 'tax_rate', $tax_rate, 'text');
        self::field('IRPF (%)', 'irpf_rate', $irpf_rate, 'text');
        self::field('Notas', 'notes', $notes, 'textarea', 'vvd-col-12');
        echo '<div class="vvd-col-12"><label class="vvd-label">Líneas</label><div class="vvd-lines">';
        foreach ($lines as $line) {
            self::line_row($line);
        }
        self::line_row(['description' => '', 'quantity' => '1', 'unit_price' => ''], true);
        echo '</div><p><button class="vvd-btn light vvd-add-line">+ Añadir línea</button></p></div>';
        echo '</div>';
        echo '<div class="vvd-totals"><p><strong>Base imponible:</strong> <span id="vvd-base">0.00 €</span></p><p><strong>IVA:</strong> <span id="vvd-tax">0.00 €</span></p><p><strong>IRPF:</strong> <span id="vvd-irpf">0.00 €</span></p><p style="font-size:20px"><strong>Total:</strong> <span id="vvd-total">0.00 €</span></p></div>';
        echo '<p><button class="vvd-btn" type="submit">Guardar documento</button></p></form></div></div>';
    }

    protected static function line_row($line, $template = false) {
        $class = $template ? 'line vvd-line-template' : 'line vvd-line';
        $style = $template ? 'style="display:none"' : '';
        echo '<div class="' . esc_attr($class) . '" ' . $style . '>';
        echo '<div><input class="vvd-input" name="description[]" placeholder="Concepto" value="' . esc_attr($line['description'] ?? '') . '"></div>';
        echo '<div><input class="vvd-input vvd-qty" name="quantity[]" placeholder="Cantidad" value="' . esc_attr($line['quantity'] ?? '1') . '"></div>';
        echo '<div><input class="vvd-input vvd-price" name="unit_price[]" placeholder="Precio unitario" value="' . esc_attr($line['unit_price'] ?? '') . '"><small>Total línea: <span class="vvd-line-total">0.00 €</span></small></div>';
        echo '<div><button class="vvd-btn light vvd-remove-line">×</button></div>';
        echo '</div>';
    }

    protected static function field($label, $name, $value = '', $type = 'text', $class = 'vvd-col-6') {
        echo '<div class="' . esc_attr($class) . '"><label class="vvd-label">' . esc_html($label) . '</label>';
        if ($type === 'textarea') {
            echo '<textarea class="vvd-textarea" name="' . esc_attr($name) . '">' . esc_textarea($value) . '</textarea>';
        } else {
            echo '<input id="' . esc_attr($name) . '" class="vvd-input" type="' . esc_attr($type) . '" name="' . esc_attr($name) . '" value="' . esc_attr($value) . '">';
        }
        echo '</div>';
    }

    public static function save_client() {
        if (!current_user_can('manage_options')) { wp_die('No autorizado'); }
        check_admin_referer('vvd_save_client');
        global $wpdb;
        $client_id = (int) ($_POST['client_id'] ?? 0);
        $data = [
            'name' => sanitize_text_field($_POST['name'] ?? ''),
            'display_name' => sanitize_text_field($_POST['display_name'] ?? ''),
            'company_name' => sanitize_text_field($_POST['company_name'] ?? ''),
            'tax_id' => sanitize_text_field($_POST['tax_id'] ?? ''),
            'contact_name' => sanitize_text_field($_POST['contact_name'] ?? ''),
            'salutation' => sanitize_text_field($_POST['salutation'] ?? ''),
            'first_name' => sanitize_text_field($_POST['first_name'] ?? ''),
            'last_name' => sanitize_text_field($_POST['last_name'] ?? ''),
            'address' => sanitize_textarea_field($_POST['address'] ?? ''),
            'billing_attention' => sanitize_text_field($_POST['billing_attention'] ?? ''),
            'billing_street2' => sanitize_text_field($_POST['billing_street2'] ?? ''),
            'city' => sanitize_text_field($_POST['city'] ?? ''),
            'state' => sanitize_text_field($_POST['state'] ?? ''),
            'postcode' => sanitize_text_field($_POST['postcode'] ?? ''),
            'country' => sanitize_text_field($_POST['country'] ?? ''),
            'shipping_attention' => sanitize_text_field($_POST['shipping_attention'] ?? ''),
            'shipping_address' => sanitize_textarea_field($_POST['shipping_address'] ?? ''),
            'shipping_street2' => sanitize_text_field($_POST['shipping_street2'] ?? ''),
            'shipping_city' => sanitize_text_field($_POST['shipping_city'] ?? ''),
            'shipping_state' => sanitize_text_field($_POST['shipping_state'] ?? ''),
            'shipping_postcode' => sanitize_text_field($_POST['shipping_postcode'] ?? ''),
            'shipping_country' => sanitize_text_field($_POST['shipping_country'] ?? ''),
            'email' => sanitize_email($_POST['email'] ?? ''),
            'phone' => sanitize_text_field($_POST['phone'] ?? ''),
            'mobile' => sanitize_text_field($_POST['mobile'] ?? ''),
            'website' => esc_url_raw($_POST['website'] ?? ''),
            'client_status' => sanitize_text_field($_POST['client_status'] ?? ''),
            'currency_code' => sanitize_text_field($_POST['currency_code'] ?? ''),
            'payment_terms' => sanitize_text_field($_POST['payment_terms'] ?? ''),
            'payment_terms_label' => sanitize_text_field($_POST['payment_terms_label'] ?? ''),
            'department' => sanitize_text_field($_POST['department'] ?? ''),
            'designation' => sanitize_text_field($_POST['designation'] ?? ''),
            'contact_type' => sanitize_text_field($_POST['contact_type'] ?? ''),
            'taxable' => sanitize_text_field($_POST['taxable'] ?? ''),
            'tax_name' => sanitize_text_field($_POST['tax_name'] ?? ''),
            'tax_percentage' => sanitize_text_field($_POST['tax_percentage'] ?? ''),
            'exemption_reason' => sanitize_textarea_field($_POST['exemption_reason'] ?? ''),
            'notes' => sanitize_textarea_field($_POST['notes'] ?? ''),
            'updated_at' => current_time('mysql'),
        ];
        if ($client_id) {
            $wpdb->update(self::table('vvd_clients'), $data, ['id' => $client_id]);
            self::log_event('client_updated', 'Cliente actualizado: ' . $data['name'], null, $client_id, $data);
        } else {
            $data['created_at'] = current_time('mysql');
            $wpdb->insert(self::table('vvd_clients'), $data);
            $client_id = (int) $wpdb->insert_id;
            self::log_event('client_created', 'Cliente creado: ' . $data['name'], null, $client_id, $data);
        }
        wp_safe_redirect(admin_url('admin.php?page=vvd_clientes'));
        exit;
    }

    public static function delete_client() {
        if (!current_user_can('manage_options')) { wp_die('No autorizado'); }
        $client_id = (int) ($_POST['client_id'] ?? 0);
        check_admin_referer('vvd_delete_client_' . $client_id);
        global $wpdb;
        $invoices_count = (int) $wpdb->get_var($wpdb->prepare('SELECT COUNT(*) FROM ' . self::table('vvd_invoices') . ' WHERE client_id=%d', $client_id));
        if ($invoices_count > 0) {
            wp_die('No se puede eliminar un cliente con documentos asociados.');
        }
        $client = self::client($client_id);
        $wpdb->delete(self::table('vvd_clients'), ['id' => $client_id]);
        self::log_event('client_deleted', 'Cliente eliminado: ' . ($client['name'] ?? ('ID ' . $client_id)), null, $client_id);
        wp_safe_redirect(admin_url('admin.php?page=vvd_clientes'));
        exit;
    }

    public static function save_settings() {
        if (!current_user_can('manage_options')) { wp_die('No autorizado'); }
        check_admin_referer('vvd_save_settings');
        global $wpdb;
        $table = self::table('vvd_settings');
        $id = (int) $wpdb->get_var("SELECT id FROM {$table} ORDER BY id ASC LIMIT 1");
        $data = [
            'issuer_name' => sanitize_text_field($_POST['issuer_name'] ?? ''),
            'issuer_tax_id' => sanitize_text_field($_POST['issuer_tax_id'] ?? ''),
            'issuer_address' => sanitize_textarea_field($_POST['issuer_address'] ?? ''),
            'issuer_email' => sanitize_email($_POST['issuer_email'] ?? ''),
            'issuer_phone' => sanitize_text_field($_POST['issuer_phone'] ?? ''),
            'issuer_iban' => sanitize_text_field($_POST['issuer_iban'] ?? ''),
            'default_series' => strtoupper(sanitize_text_field($_POST['default_series'] ?? 'A')),
            'rectificative_series' => strtoupper(sanitize_text_field($_POST['rectificative_series'] ?? 'RT')),
            'quote_series' => strtoupper(sanitize_text_field($_POST['quote_series'] ?? 'P')),
            'series_padding' => max(3, (int) ($_POST['series_padding'] ?? 6)),
            'default_tax_rate' => (float) str_replace(',', '.', $_POST['default_tax_rate'] ?? 21),
            'default_irpf_rate' => (float) str_replace(',', '.', $_POST['default_irpf_rate'] ?? 15),
            'logo_url' => esc_url_raw($_POST['logo_url'] ?? ''),
            'background_image_url' => esc_url_raw($_POST['background_image_url'] ?? ''),
            'brand_primary' => sanitize_text_field($_POST['brand_primary'] ?? '#6997c1'),
            'brand_secondary' => sanitize_text_field($_POST['brand_secondary'] ?? '#8d8e8e'),
            'brand_background' => sanitize_text_field($_POST['brand_background'] ?? '#FFFFFF'),
            'updated_at' => current_time('mysql'),
        ];
        $wpdb->update($table, $data, ['id' => $id]);
        self::log_event('settings_updated', 'Ajustes actualizados', null, null, $data);
        wp_safe_redirect(admin_url('admin.php?page=vvd_ajustes'));
        exit;
    }

    public static function save_invoice() {
        if (!current_user_can('manage_options')) { wp_die('No autorizado'); }
        check_admin_referer('vvd_save_invoice');
        global $wpdb;
        $settings = self::settings();
        $invoice_id = (int) ($_POST['invoice_id'] ?? 0);
        $type = sanitize_text_field($_POST['invoice_type'] ?? 'standard');
        $descs = $_POST['description'] ?? [];
        $qtys = $_POST['quantity'] ?? [];
        $prices = $_POST['unit_price'] ?? [];
        $lines = [];
        $base = 0.0;
        foreach ($descs as $i => $desc) {
            $desc = sanitize_text_field($desc);
            if ($desc === '') { continue; }
            $qty = (float) str_replace(',', '.', $qtys[$i] ?? 0);
            $price = (float) str_replace(',', '.', $prices[$i] ?? 0);
            $line_total = round($qty * $price, 2);
            $lines[] = [
                'description' => $desc,
                'quantity' => $qty,
                'unit_price' => $price,
                'line_total' => $line_total,
            ];
            $base += $line_total;
        }
        if (!$lines) {
            wp_die('Debes añadir al menos una línea');
        }

        $tax_rate = ($type === 'quote') ? 0.0 : (float) str_replace(',', '.', $_POST['tax_rate'] ?? 21);
        $irpf_rate = ($type === 'quote') ? 0.0 : (float) str_replace(',', '.', $_POST['irpf_rate'] ?? 0);
        $tax_amount = round($base * $tax_rate / 100, 2);
        $irpf_amount = round($base * $irpf_rate / 100, 2);
        $total = round($base + $tax_amount - $irpf_amount, 2);
        $client_id = (int) ($_POST['client_id'] ?? 0);
        $issue_date = sanitize_text_field($_POST['issue_date'] ?? current_time('Y-m-d'));
        $original_invoice_id = !empty($_POST['original_invoice_id']) ? (int) $_POST['original_invoice_id'] : null;
        $notes = sanitize_textarea_field($_POST['notes'] ?? '');
        $client_row = $client_id ? self::client($client_id) : [];
        $invoice_snapshot = [
            'source_origin' => 'plugin_generated',
            'billing_name' => sanitize_text_field($client_row['name'] ?? ''),
            'billing_address' => sanitize_text_field($client_row['address'] ?? ''),
            'billing_city' => sanitize_text_field($client_row['city'] ?? ''),
            'billing_state' => sanitize_text_field($client_row['state'] ?? ''),
            'billing_postcode' => sanitize_text_field($client_row['postcode'] ?? ''),
            'billing_country' => sanitize_text_field($client_row['country'] ?? ''),
            'client_email' => sanitize_email($client_row['email'] ?? ''),
            'client_phone' => sanitize_text_field($client_row['phone'] ?? ''),
            'client_mobile' => sanitize_text_field($client_row['mobile'] ?? ''),
            'client_external_id' => sanitize_text_field($client_row['external_id'] ?? ''),
            'subtotal' => $base,
            'discount_total' => 0,
            'withholding_total' => $irpf_amount,
            'balance_due' => $total,
            'currency_code' => 'EUR',
        ];

        if ($invoice_id) {
            $existing = self::invoice($invoice_id);
            $hash_prev = ($type === 'quote') ? '' : self::latest_hash_before();
            $hash_current = self::compute_invoice_hash([
                'series' => $existing['series'], 'number' => $existing['number'], 'issue_date' => $issue_date, 'client_id' => $client_id,
                'taxable_base' => $base, 'tax_amount' => $tax_amount, 'irpf_amount' => $irpf_amount, 'total_amount' => $total,
            ], $lines, $hash_prev);
            $wpdb->update(self::table('vvd_invoices'), [
                'invoice_type' => $type,
                'issue_date' => $issue_date,
                'client_id' => $client_id,
                'original_invoice_id' => $original_invoice_id,
                'notes' => $notes,
                'terms_conditions' => sanitize_textarea_field($_POST['terms_conditions'] ?? ($existing['terms_conditions'] ?? '')),
                'taxable_base' => $base,
                'tax_rate' => $tax_rate,
                'tax_amount' => $tax_amount,
                'irpf_rate' => $irpf_rate,
                'irpf_amount' => $irpf_amount,
                'total_amount' => $total,
                'status' => ($type === 'quote' ? 'quote' : 'issued'),
                'source_origin' => 'plugin_generated',
                'client_external_id' => $invoice_snapshot['client_external_id'],
                'billing_name' => $invoice_snapshot['billing_name'],
                'billing_address' => $invoice_snapshot['billing_address'],
                'billing_city' => $invoice_snapshot['billing_city'],
                'billing_state' => $invoice_snapshot['billing_state'],
                'billing_postcode' => $invoice_snapshot['billing_postcode'],
                'billing_country' => $invoice_snapshot['billing_country'],
                'client_email' => $invoice_snapshot['client_email'],
                'client_phone' => $invoice_snapshot['client_phone'],
                'client_mobile' => $invoice_snapshot['client_mobile'],
                'currency_code' => $invoice_snapshot['currency_code'],
                'subtotal' => $invoice_snapshot['subtotal'],
                'discount_total' => $invoice_snapshot['discount_total'],
                'withholding_total' => $invoice_snapshot['withholding_total'],
                'balance_due' => $invoice_snapshot['balance_due'],
                'verifactu_status' => ($type === 'quote' ? 'not-applicable' : 'draft-ready'),
                'hash_prev' => $hash_prev,
                'hash_current' => $hash_current,
                'updated_at' => current_time('mysql'),
            ], ['id' => $invoice_id]);
            $wpdb->delete(self::table('vvd_invoice_lines'), ['invoice_id' => $invoice_id]);
            $sort_order = 0;
            foreach ($lines as $line) {
                $wpdb->insert(self::table('vvd_invoice_lines'), ['invoice_id' => $invoice_id, 'item_name' => $line['description'], 'item_desc' => $line['description'], 'usage_unit' => '', 'discount_amount' => 0, 'sort_order' => $sort_order++] + $line);
            }
            self::log_event('invoice_updated', 'Documento actualizado: ' . ($existing['code'] ?: $invoice_id), $invoice_id, $client_id, ['total' => $total]);
        } else {
            $series_data = self::reserve_series_number($type);
            $hash_prev = ($type === 'quote') ? '' : self::latest_hash_before();
            $hash_current = self::compute_invoice_hash([
                'series' => $series_data['series'], 'number' => $series_data['number'], 'issue_date' => $issue_date, 'client_id' => $client_id,
                'taxable_base' => $base, 'tax_amount' => $tax_amount, 'irpf_amount' => $irpf_amount, 'total_amount' => $total,
            ], $lines, $hash_prev);
            $wpdb->insert(self::table('vvd_invoices'), [
                'series' => $series_data['series'],
                'series_label' => $series_data['series_label'],
                'number' => $series_data['number'],
                'code' => $series_data['code'],
                'invoice_type' => $type,
                'issue_date' => $issue_date,
                'client_id' => $client_id,
                'original_invoice_id' => $original_invoice_id,
                'notes' => $notes,
                'terms_conditions' => sanitize_textarea_field($_POST['terms_conditions'] ?? ''),
                'taxable_base' => $base,
                'tax_rate' => $tax_rate,
                'tax_amount' => $tax_amount,
                'irpf_rate' => $irpf_rate,
                'irpf_amount' => $irpf_amount,
                'total_amount' => $total,
                'status' => ($type === 'quote' ? 'quote' : 'issued'),
                'source_origin' => 'plugin_generated',
                'client_external_id' => $invoice_snapshot['client_external_id'],
                'billing_name' => $invoice_snapshot['billing_name'],
                'billing_address' => $invoice_snapshot['billing_address'],
                'billing_city' => $invoice_snapshot['billing_city'],
                'billing_state' => $invoice_snapshot['billing_state'],
                'billing_postcode' => $invoice_snapshot['billing_postcode'],
                'billing_country' => $invoice_snapshot['billing_country'],
                'client_email' => $invoice_snapshot['client_email'],
                'client_phone' => $invoice_snapshot['client_phone'],
                'client_mobile' => $invoice_snapshot['client_mobile'],
                'currency_code' => $invoice_snapshot['currency_code'],
                'subtotal' => $invoice_snapshot['subtotal'],
                'discount_total' => $invoice_snapshot['discount_total'],
                'withholding_total' => $invoice_snapshot['withholding_total'],
                'balance_due' => $invoice_snapshot['balance_due'],
                'verifactu_status' => ($type === 'quote' ? 'not-applicable' : 'draft-ready'),
                'hash_prev' => $hash_prev,
                'hash_current' => $hash_current,
                'created_at' => current_time('mysql'),
                'updated_at' => current_time('mysql'),
            ]);
            $invoice_id = (int) $wpdb->insert_id;
            $sort_order = 0;
            foreach ($lines as $line) {
                $wpdb->insert(self::table('vvd_invoice_lines'), ['invoice_id' => $invoice_id, 'item_name' => $line['description'], 'item_desc' => $line['description'], 'usage_unit' => '', 'discount_amount' => 0, 'sort_order' => $sort_order++] + $line);
            }
            self::log_event('invoice_created', 'Documento creado: ' . $series_data['code'], $invoice_id, $client_id, ['total' => $total, 'type' => $type]);
        }

        self::maybe_generate_pdf($invoice_id, true);
        wp_safe_redirect(admin_url('admin.php?page=vvd_facturas&view=' . $invoice_id));
        exit;
    }

    protected static function latest_hash_before() {
        global $wpdb;
        $hash = $wpdb->get_var('SELECT hash_current FROM ' . self::table('vvd_invoices') . ' WHERE invoice_type != "quote" AND hash_current IS NOT NULL AND hash_current != "" ORDER BY id DESC LIMIT 1');
        return $hash ?: '';
    }

    protected static function compute_invoice_hash($invoice, $lines, $prev_hash = '') {
        $payload = [
            'series' => $invoice['series'] ?? '',
            'number' => $invoice['number'] ?? 0,
            'date' => $invoice['issue_date'] ?? '',
            'client' => $invoice['client_id'] ?? 0,
            'base' => $invoice['taxable_base'] ?? 0,
            'tax' => $invoice['tax_amount'] ?? 0,
            'irpf' => $invoice['irpf_amount'] ?? 0,
            'total' => $invoice['total_amount'] ?? 0,
            'prev' => $prev_hash,
            'lines' => array_map(function($line){
                return [
                    'd' => $line['description'],
                    'q' => (float) $line['quantity'],
                    'u' => (float) $line['unit_price'],
                    't' => (float) $line['line_total'],
                ];
            }, $lines),
        ];
        return hash('sha256', wp_json_encode($payload));
    }

    protected static function uploads_dir() {
        $upload = wp_upload_dir();
        $dir = trailingslashit($upload['basedir']) . 'vvd-invoices/';
        if (!file_exists($dir)) {
            wp_mkdir_p($dir);
        }
        return $dir;
    }

    protected static function find_chromium_binary() {
        $candidates = [
            defined('VVD_CHROMIUM_PATH') ? VVD_CHROMIUM_PATH : '',
            '/usr/bin/chromium',
            '/usr/bin/chromium-browser',
            '/usr/bin/google-chrome',
            '/usr/bin/google-chrome-stable',
            '/snap/bin/chromium',
            'chromium',
            'chromium-browser',
            'google-chrome',
        ];

        foreach ($candidates as $candidate) {
            $candidate = trim((string) $candidate);
            if ($candidate === '') {
                continue;
            }
            if ($candidate[0] === '/' && @is_executable($candidate)) {
                return $candidate;
            }
            if (strpos($candidate, '/') === false && function_exists('shell_exec')) {
                $resolved = @shell_exec('command -v ' . escapeshellarg($candidate) . ' 2>/dev/null');
                $resolved = trim((string) $resolved);
                if ($resolved !== '' && @is_executable($resolved)) {
                    return $resolved;
                }
            }
        }

        return '';
    }

    protected static function generate_pdf_from_html($html, $path) {
        $binary = self::find_chromium_binary();
        if ($binary === '' || !function_exists('shell_exec')) {
            return false;
        }

        $tmp_dir = trailingslashit(self::uploads_dir()) . 'tmp/';
        if (!file_exists($tmp_dir)) {
            wp_mkdir_p($tmp_dir);
        }

        $tmp_html = tempnam($tmp_dir, 'vvd_html_');
        if (!$tmp_html) {
            return false;
        }
        $tmp_html_file = $tmp_html . '.html';
        @rename($tmp_html, $tmp_html_file);
        file_put_contents($tmp_html_file, $html);

        $cmd = escapeshellarg($binary)
            . ' --headless=new --disable-gpu --no-sandbox --allow-file-access-from-files'
            . ' --print-to-pdf=' . escapeshellarg($path)
            . ' ' . escapeshellarg('file://' . $tmp_html_file)
            . ' 2>&1';

        @shell_exec($cmd);
        @unlink($tmp_html_file);

        return file_exists($path) && filesize($path) > 1000;
    }

    protected static function maybe_generate_pdf($invoice_id, $save = true) {
        global $wpdb;
        $invoice = self::invoice($invoice_id);
        $client = self::client($invoice['client_id']);
        $settings = self::settings();
        $filename = 'document_' . preg_replace('/[^A-Za-z0-9\-_]/', '_', ($invoice['code'] ?: self::invoice_code($invoice['series'], $invoice['number'], (int) ($settings['series_padding'] ?: 6)))) . '.pdf';
        $path = self::uploads_dir() . $filename;

        $html = self::render_document_html($invoice, self::invoice_lines($invoice_id), $client, $settings, false);
        $generated = self::generate_pdf_from_html($html, $path);
        $engine = 'html_print';

        if (!$generated) {
            $pdf = new VVD_PDF();
            $binary = $pdf->render($invoice, self::invoice_lines($invoice_id), $client, $settings);
            if ($save) {
                file_put_contents($path, $binary);
            }
            $engine = 'legacy_fallback';
        } else {
            $binary = file_get_contents($path);
        }

        if ($save && file_exists($path)) {
            $wpdb->update(self::table('vvd_invoices'), ['pdf_path' => $path], ['id' => $invoice_id]);
            self::log_event('pdf_generated', 'PDF generado: ' . $filename, $invoice_id, $invoice['client_id'], ['engine' => $engine]);
        }

        return ['binary' => $binary, 'filename' => $filename, 'path' => $path, 'engine' => $engine];
    }


    protected static function document_title($invoice) {
        $type = (string) ($invoice['invoice_type'] ?? '');
        if ($type === 'rectificativa') {
            return 'FACTURA RECTIFICATIVA';
        }
        if ($type === 'quote') {
            return 'PRESUPUESTO';
        }
        return 'FACTURA';
    }

    protected static function fmt_money_html($amount) {
        return number_format((float) $amount, 2, ',', '.') . ' €';
    }

    protected static function render_document_html($invoice, $lines, $client, $settings, $auto_print = false) {
        $title = self::document_title($invoice);
        $code = $invoice['code'] ?: self::invoice_code($invoice['series'], $invoice['number'], (int) ($settings['series_padding'] ?: 6));
        $issue_date = !empty($invoice['issue_date']) ? date_i18n('d/m/Y', strtotime($invoice['issue_date'])) : '-';
        $due_date = !empty($invoice['due_date']) ? date_i18n('d/m/Y', strtotime($invoice['due_date'])) : '-';
        $status = self::invoice_status_label($invoice['status'] ?: 'issued');
        $primary = $settings['brand_primary'] ?: '#6997c1';
        $secondary = $settings['brand_secondary'] ?: '#8d8e8e';
        $logo = $settings['logo_url'] ?: 'https://vexeldot.es/wp-content/uploads/cropped-Logo-VexelDot.png';
        $issuer_name = $settings['issuer_name'] ?: get_bloginfo('name');
        $issuer_tax_id = $settings['issuer_tax_id'] ?: '-';
        $issuer_address = $settings['issuer_address'] ?: '-';
        $issuer_email = $settings['issuer_email'] ?: '';
        $issuer_phone = $settings['issuer_phone'] ?: '';
        $issuer_iban = $settings['issuer_iban'] ?: '';
        $notes = trim((string) ($invoice['notes'] ?: 'Gracias por confiar en VexelDot.'));
        $tax_label = (($invoice['invoice_type'] ?? '') === 'quote') ? 'Impuestos' : ('IVA ' . number_format((float) ($invoice['tax_rate'] ?? 0), 0, ',', '.') . '%');
        $rows_html = '';

        foreach ($lines as $line) {
            $desc = nl2br(esc_html((string) ($line['description'] ?? '')));
            $qty = rtrim(rtrim(number_format((float) ($line['quantity'] ?? 0), 2, ',', '.'), '0'), ',');
            if ($qty === '') { $qty = '0'; }
            $unit = self::fmt_money_html((float) ($line['unit_price'] ?? 0));
            $tax = (($invoice['invoice_type'] ?? '') === 'quote') ? '0%' : (number_format((float) ($invoice['tax_rate'] ?? 0), 0, ',', '.') . '%');
            $total = self::fmt_money_html((float) ($line['line_total'] ?? 0));
            $parts = preg_split('/\r\n|\r|\n/', trim((string) ($line['description'] ?? '')));
            $headline = esc_html((string) ($parts[0] ?? ''));
            $body = trim(implode("\n", array_slice($parts, 1)));
            $desc_html = '<strong>' . ($headline !== '' ? $headline : 'Concepto') . '</strong>';
            if ($body !== '') {
                $desc_html .= '<br />' . nl2br(esc_html($body));
            }
            $rows_html .= '<tr>'
                . '<td class="col-desc">' . $desc_html . '</td>'
                . '<td class="col-qty">' . esc_html($qty) . '</td>'
                . '<td class="col-price">' . esc_html($unit) . '</td>'
                . '<td class="col-tax">' . esc_html($tax) . '</td>'
                . '<td class="col-total">' . esc_html($total) . '</td>'
                . '</tr>';
        }

        if ($rows_html === '') {
            $rows_html = '<tr><td class="col-desc"><strong>Sin líneas</strong></td><td class="col-qty">0</td><td class="col-price">0,00 €</td><td class="col-tax">0%</td><td class="col-total">0,00 €</td></tr>';
        }

        $base = self::fmt_money_html((float) ($invoice['taxable_base'] ?? 0));
        $tax_amount = self::fmt_money_html((float) ($invoice['tax_amount'] ?? 0));
        $irpf_amount = (float) ($invoice['irpf_amount'] ?? 0);
        $total = self::fmt_money_html((float) ($invoice['total_amount'] ?? 0));

        ob_start();
        ?>
<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title><?php echo esc_html($title . ' ' . $code); ?></title>
  <style>
    :root {
      --vx-primary: <?php echo esc_html($primary); ?>;
      --vx-white: #ffffff;
      --vx-gray: <?php echo esc_html($secondary); ?>;
      --vx-text: #2f3a45;
      --vx-border: #d9e1e8;
      --vx-soft: #f5f8fb;
      --vx-page-width: 210mm;
      --vx-page-min-height: 297mm;
      --vx-shadow: 0 10px 30px rgba(35, 52, 70, 0.12);
    }
    * { box-sizing: border-box; }
    body {
      margin: 0;
      padding: 32px;
      background: #eef3f7;
      font-family: Arial, Helvetica, sans-serif;
      color: var(--vx-text);
    }
    .pdf-page {
      position: relative;
      width: 100%;
      max-width: var(--vx-page-width);
      min-height: var(--vx-page-min-height);
      margin: 0 auto;
      background: var(--vx-white);
      box-shadow: var(--vx-shadow);
      overflow: hidden;
    }
    .pdf-bg {
      position: absolute;
      inset: 0;
      background: #ffffff;
      z-index: 0;
    }
    .pdf-overlay {
      position: absolute;
      inset: 0;
      background: linear-gradient(180deg, rgba(255,255,255,0.88) 0%, rgba(255,255,255,0.94) 28%, rgba(255,255,255,0.98) 100%);
      z-index: 0;
      pointer-events: none;
    }
    .pdf-content {
      position: relative;
      z-index: 1;
      padding: 18mm 14mm 10mm;
    }
    .top-accent {
      position: absolute;
      top: 0; left: 0; right: 0;
      height: 10px;
      background: linear-gradient(90deg, var(--vx-primary), #83add2 55%, var(--vx-gray));
      z-index: 2;
    }
    .header {
      display: flex;
      justify-content: space-between;
      align-items: flex-start;
      gap: 10px;
      margin-bottom: 12px;
    }
    .brand img {
      width: 205px;
      max-width: 100%;
      height: auto;
      display: block;
    }
    .doc-box {
      min-width: 210px;
      max-width: 250px;
      background: rgba(255,255,255,0.82);
      border: 1px solid var(--vx-border);
      border-radius: 12px;
      padding: 12px;
    }
    .doc-label {
      margin: 0 0 6px;
      font-size: 10px;
      letter-spacing: 0.5px;
      text-transform: uppercase;
      color: var(--vx-gray);
      font-weight: 700;
    }
    .doc-title {
      margin: 0 0 8px;
      font-size: 22px;
      line-height: 1;
      color: var(--vx-primary);
      font-weight: 800;
    }
    .meta-grid { display: grid; grid-template-columns: 1fr; gap: 5px; font-size: 11px; }
    .meta-row {
      display: flex;
      justify-content: space-between;
      gap: 10px;
      padding-bottom: 4px;
      border-bottom: 1px dashed var(--vx-border);
    }
    .meta-row strong { color: var(--vx-text); }
    .parties {
      display: grid;
      grid-template-columns: 1fr 1fr;
      gap: 10px;
      margin-bottom: 12px;
    }
    .card {
      background: rgba(255,255,255,0.9);
      border: 1px solid var(--vx-border);
      border-radius: 12px;
      padding: 12px;
    }
    .card h3 {
      margin: 0 0 7px;
      font-size: 10.5px;
      letter-spacing: 0.7px;
      text-transform: uppercase;
      color: var(--vx-primary);
    }
    .card p {
      margin: 2px 0;
      font-size: 11px;
      line-height: 1.45;
    }
    .items-wrap { margin-bottom: 10px; }
    table {
      width: 100%;
      border-collapse: collapse;
      table-layout: fixed;
      background: rgba(255,255,255,0.92);
      border: 1px solid var(--vx-border);
      border-radius: 12px;
      overflow: hidden;
    }
    thead th {
      background: var(--vx-primary);
      color: var(--vx-white);
      font-size: 10px;
      text-transform: uppercase;
      letter-spacing: 0.5px;
      padding: 7px 6px;
      text-align: left;
    }
    tbody td {
      padding: 7px 6px;
      border-top: 1px solid var(--vx-border);
      font-size: 10.5px;
      line-height: 1.35;
      vertical-align: top;
      word-break: break-word;
    }
    .col-desc { width: 56%; }
    .col-qty { width: 8%; text-align: center; }
    .col-price { width: 12%; text-align: right; }
    .col-tax { width: 9%; text-align: center; }
    .col-total { width: 15%; text-align: right; }
    tbody td.col-qty, tbody td.col-price, tbody td.col-tax, tbody td.col-total { white-space: nowrap; }
    .totals-area {
      display: grid;
      grid-template-columns: 1.35fr 0.65fr;
      gap: 10px;
      align-items: start;
      margin-bottom: 10px;
    }
    .notes { min-height: 100%; }
    .notes h4, .totals h4 {
      margin: 0 0 10px;
      color: var(--vx-primary);
      font-size: 10.5px;
      text-transform: uppercase;
      letter-spacing: 0.7px;
    }
    .notes p {
      margin: 0;
      font-size: 10.5px;
      line-height: 1.3;
    }
    .totals table { border-radius: 16px; }
    .totals td {
      padding: 6px 8px;
      font-size: 10.5px;
    }
    .totals td:last-child {
      text-align: right;
      white-space: nowrap;
      font-weight: 700;
    }
    .grand-total td {
      background: var(--vx-soft);
      font-size: 14px;
      font-weight: 800;
      color: var(--vx-primary);
    }
    .footer {
      display: flex;
      justify-content: space-between;
      gap: 10px;
      align-items: flex-end;
      border-top: 2px solid var(--vx-border);
      padding-top: 8px;
      font-size: 10px;
      color: #5f6c77;
    }
    .footer strong { color: var(--vx-text); }
    .watermark {
      position: absolute;
      right: 14mm;
      bottom: 8mm;
      font-size: 30px;
      font-weight: 800;
      color: rgba(105, 151, 193, 0.08);
      letter-spacing: 2px;
      z-index: 0;
      user-select: none;
    }
    .print-actions {
      max-width: var(--vx-page-width);
      margin: 0 auto 18px;
      display: flex;
      gap: 10px;
    }
    .print-actions button, .print-actions a {
      appearance: none;
      border: 0;
      background: var(--vx-primary);
      color: #fff;
      padding: 10px 14px;
      border-radius: 10px;
      text-decoration: none;
      font-weight: 700;
      cursor: pointer;
      font-size: 11px;
    }
    .print-actions a { background: #8d8e8e; }

    tbody td.col-desc strong { display:block; margin-bottom:2px; font-size:10.5px; }
    .card.notes, .card.totals { padding: 10px 12px; }
    .notes h4, .totals h4 { margin: 0 0 6px; font-size: 10px; }

    @media print {
      body { background: white; padding: 0; }
      .pdf-page { box-shadow: none; max-width: none; }
      .print-actions { display: none !important; }
    }
    @page { size: A4; margin: 0; }
  </style>
</head>
<body>
  <div class="print-actions">
    <button onclick="window.print()">Guardar / Imprimir PDF</button>
    <a href="<?php echo esc_url(admin_url('admin.php?page=vvd_facturas&view=' . (int) $invoice['id'])); ?>">Volver</a>
  </div>
  <div class="pdf-page">
    <div class="top-accent"></div>
    <div class="pdf-bg"></div>
    <div class="pdf-overlay"></div>
    <div class="watermark">VEXELDOT</div>

    <div class="pdf-content">
      <header class="header">
        <div class="brand">
          <img src="<?php echo esc_url($logo); ?>" alt="VexelDot" />
        </div>

        <div class="doc-box">
          <p class="doc-label">Documento</p>
          <p class="doc-title"><?php echo esc_html($title); ?></p>

          <div class="meta-grid">
            <div class="meta-row"><span>Número</span><strong><?php echo esc_html($code); ?></strong></div>
            <div class="meta-row"><span>Fecha</span><strong><?php echo esc_html($issue_date); ?></strong></div>
            <div class="meta-row"><span>Vencimiento</span><strong><?php echo esc_html($due_date); ?></strong></div>
            <div class="meta-row"><span>Estado</span><strong><?php echo esc_html($status); ?></strong></div>
          </div>
        </div>
      </header>

      <section class="parties">
        <div class="card">
          <h3>Emisor</h3>
          <p><strong><?php echo esc_html($issuer_name); ?></strong></p>
          <p>CIF: <?php echo esc_html($issuer_tax_id); ?></p>
          <p><?php echo nl2br(esc_html($issuer_address)); ?></p>
          <?php if ($issuer_email) : ?><p><?php echo esc_html($issuer_email); ?></p><?php endif; ?>
          <?php if ($issuer_phone) : ?><p><?php echo esc_html($issuer_phone); ?></p><?php endif; ?>
        </div>

        <div class="card">
          <h3>Cliente</h3>
          <p><strong><?php echo esc_html($client['name'] ?: '-'); ?></strong></p>
          <p>CIF: <?php echo esc_html($client['tax_id'] ?: '-'); ?></p>
          <p><?php echo nl2br(esc_html($client['address'] ?: '-')); ?></p>
          <?php if (!empty($client['email'])) : ?><p><?php echo esc_html($client['email']); ?></p><?php endif; ?>
          <?php if (!empty($client['phone'])) : ?><p><?php echo esc_html($client['phone']); ?></p><?php endif; ?>
        </div>
      </section>

      <section class="items-wrap">
        <table>
          <thead>
            <tr>
              <th class="col-desc">Concepto</th>
              <th class="col-qty">Cant.</th>
              <th class="col-price">Precio</th>
              <th class="col-tax">Impuesto</th>
              <th class="col-total">Importe</th>
            </tr>
          </thead>
          <tbody>
            <?php echo $rows_html; ?>
          </tbody>
        </table>
      </section>

      <section class="totals-area">
        <div class="card notes">
          <h4>Notas</h4>
          <p><?php echo nl2br(esc_html($notes)); ?></p>
        </div>

        <div class="card totals">
          <h4>Resumen</h4>
          <table>
            <tbody>
              <tr>
                <td>Base imponible</td>
                <td><?php echo esc_html($base); ?></td>
              </tr>
              <?php if (($invoice['invoice_type'] ?? '') !== 'quote') : ?>
              <tr>
                <td><?php echo esc_html($tax_label); ?></td>
                <td><?php echo esc_html($tax_amount); ?></td>
              </tr>
              <?php if ((float) ($invoice['irpf_rate'] ?? 0) > 0) : ?>
              <tr>
                <td>IRPF <?php echo esc_html(number_format((float) ($invoice['irpf_rate'] ?? 0), 0, ',', '.')); ?>%</td>
                <td><?php echo esc_html(($irpf_amount < 0 ? '-' : '') . self::fmt_money_html(abs($irpf_amount))); ?></td>
              </tr>
              <?php endif; ?>
              <?php endif; ?>
              <tr class="grand-total">
                <td>Total</td>
                <td><?php echo esc_html($total); ?></td>
              </tr>
            </tbody>
          </table>
        </div>
      </section>

      <footer class="footer">
        <div>
          <strong>Forma de pago:</strong> Transferencia bancaria<br />
          <?php if ($issuer_iban) : ?><strong>IBAN:</strong> <?php echo esc_html($issuer_iban); ?><?php endif; ?>
        </div>
        <div>
          <strong>Referencia:</strong> <?php echo esc_html($code); ?>
        </div>
      </footer>
    </div>
  </div>
  <?php if ($auto_print) : ?>
  <script>window.addEventListener('load', function(){ setTimeout(function(){ window.print(); }, 250); });</script>
  <?php endif; ?>
</body>
</html>
        <?php
        return ob_get_clean();
    }

    public static function print_document() {
        if (!current_user_can('manage_options')) { wp_die('No autorizado'); }
        $invoice_id = (int) ($_GET['invoice_id'] ?? 0);
        check_admin_referer('vvd_print_' . $invoice_id);
        $invoice = self::invoice($invoice_id);
        if (empty($invoice)) { wp_die('Documento no encontrado'); }
        $client = self::client($invoice['client_id']);
        $settings = self::settings();
        nocache_headers();
        echo self::render_document_html($invoice, self::invoice_lines($invoice_id), $client, $settings, true);
        exit;
    }


    public static function download_pdf() {
        if (!current_user_can('manage_options')) { wp_die('No autorizado'); }
        $invoice_id = (int) ($_GET['invoice_id'] ?? 0);
        check_admin_referer('vvd_pdf_' . $invoice_id);
        $invoice = self::invoice($invoice_id);
        if (empty($invoice)) { wp_die('Documento no encontrado'); }
        $client = self::client($invoice['client_id']);
        $settings = self::settings();
        nocache_headers();
        echo self::render_document_html($invoice, self::invoice_lines($invoice_id), $client, $settings, true);
        exit;
    }

    public static function public_document() {
        $invoice_id = (int) ($_GET['invoice_id'] ?? 0);
        $token = sanitize_text_field((string) ($_GET['token'] ?? ''));
        $invoice = self::invoice($invoice_id);
        if (empty($invoice) || !self::verify_public_document_token($invoice, $token)) {
            wp_die('Documento no disponible');
        }
        $client = self::client($invoice['client_id']);
        $settings = self::settings();
        nocache_headers();
        echo self::render_document_html($invoice, self::invoice_lines($invoice_id), $client, $settings, true);
        exit;
    }


    protected static function invoice_status_label($status) {
        $status = strtolower(trim((string) $status));
        $labels = [
            'issued' => 'Emitida',
            'sent' => 'Enviada',
            'quote' => 'Presupuesto',
            'draft' => 'Borrador',
            'paid' => 'Pagada',
            'cancelled' => 'Cancelada',
        ];
        return $labels[$status] ?? ucfirst((string) $status ?: '-');
    }

    protected static function email_html($invoice, $client, $settings, $download_url) {
        $code = $invoice['code'] ?: self::invoice_code($invoice['series'], $invoice['number'], (int) ($settings['series_padding'] ?: 6));
        $primary = $settings['brand_primary'] ?: '#6997c1';
        $secondary = $settings['brand_secondary'] ?: '#8d8e8e';
        $bg = $settings['brand_background'] ?: '#ffffff';
        $logo = $settings['logo_url'] ?: '';
        $title = ($invoice['invoice_type'] === 'quote') ? 'Tu presupuesto está listo' : 'Tu factura está lista';
        $button_label = ($invoice['invoice_type'] === 'quote') ? 'Descargar presupuesto' : 'Descargar factura';
        $intro = ($invoice['invoice_type'] === 'quote')
            ? 'Hola ' . esc_html($client['name']) . ', aquí tienes tu presupuesto.'
            : 'Hola ' . esc_html($client['name']) . ', aquí tienes tu factura.';
        return '<!doctype html><html><body style="margin:0;padding:0;background:' . esc_attr($bg) . ';font-family:Arial,Helvetica,sans-serif;color:#2f3a45">'
            . '<div style="max-width:760px;margin:0 auto;padding:32px 16px">'
            . '<div style="background:#ffffff;border:1px solid #d9e1e8;border-radius:24px;overflow:hidden;box-shadow:0 20px 50px rgba(0,0,0,.08)">'
            . '<div style="height:16px;background:linear-gradient(90deg,' . esc_attr($primary) . ', #83add2 55%, ' . esc_attr($secondary) . ');"></div>'
            . '<div style="padding:30px 28px 18px 28px">'
            . '<table role="presentation" width="100%" cellspacing="0" cellpadding="0" style="border-collapse:collapse"><tr>'
            . '<td valign="top" style="padding:0 12px 18px 0">'
            . ($logo ? '<img src="' . esc_url($logo) . '" alt="Logo" style="width:240px;max-width:100%;height:auto;display:block">' : '<div style="font-size:28px;font-weight:700;color:' . esc_attr($primary) . '">' . esc_html($settings['issuer_name']) . '</div>')
            . '</td>'
            . '<td valign="top" align="right" style="padding:0 0 18px 12px">'
            . '<div style="display:inline-block;min-width:250px;max-width:300px;text-align:left;background:#ffffff;border:1px solid #d9e1e8;border-radius:16px;padding:18px">'
            . '<p style="margin:0 0 6px 0;font-size:12px;letter-spacing:1.2px;text-transform:uppercase;color:' . esc_attr($secondary) . ';font-weight:700">Documento</p>'
            . '<p style="margin:0 0 14px 0;font-size:28px;line-height:1;color:' . esc_attr($primary) . ';font-weight:800">' . esc_html(strtoupper(self::doc_label($invoice['invoice_type']))) . '</p>'
            . '<p style="margin:0 0 8px 0;padding-bottom:6px;border-bottom:1px dashed #d9e1e8;font-size:13px"><span style="color:#6b7280">Número</span><span style="float:right;font-weight:700;color:#2f3a45">' . esc_html($code) . '</span></p>'
            . '<p style="margin:0 0 8px 0;padding-bottom:6px;border-bottom:1px dashed #d9e1e8;font-size:13px"><span style="color:#6b7280">Fecha</span><span style="float:right;font-weight:700;color:#2f3a45">' . esc_html(date_i18n('d/m/Y', strtotime($invoice['issue_date']))) . '</span></p>'
            . '<p style="margin:0;font-size:13px"><span style="color:#6b7280">Estado</span><span style="float:right;font-weight:700;color:#2f3a45">' . esc_html(self::invoice_status_label($invoice['status'] ?: 'issued')) . '</span></p>'
            . '</div>'
            . '</td></tr></table>'
            . '<p style="margin:0 0 20px 0;font-size:16px;line-height:1.6;color:#5f6c77">' . $intro . '</p>'
            . '<div style="background:#ffffff;border:1px solid #d9e1e8;border-radius:18px;padding:20px 18px;margin:0 0 24px 0">'
            . '<p style="margin:0 0 10px 0;font-size:14px"><strong>Cliente:</strong> ' . esc_html($client['name']) . '</p>'
            . '<p style="margin:0 0 10px 0;font-size:14px"><strong>Base imponible:</strong> ' . esc_html(number_format((float) $invoice['taxable_base'], 2, ',', '.')) . ' €</p>'
            . (($invoice['invoice_type'] !== 'quote') ? '<p style="margin:0 0 10px 0;font-size:14px"><strong>IVA:</strong> ' . esc_html(number_format((float) $invoice['tax_amount'], 2, ',', '.')) . ' €</p>' : '')
            . (($invoice['invoice_type'] !== 'quote' && (float) $invoice['irpf_rate'] > 0) ? '<p style="margin:0 0 10px 0;font-size:14px"><strong>IRPF:</strong> ' . esc_html(number_format((float) $invoice['irpf_amount'], 2, ',', '.')) . ' €</p>' : '')
            . '<p style="margin:12px 0 0 0;font-size:24px;color:' . esc_attr($primary) . ';font-weight:800"><strong>Total:</strong> ' . esc_html(number_format((float) $invoice['total_amount'], 2, ',', '.')) . ' €</p>'
            . '</div>'
            . '<div style="text-align:center;padding:8px 0 6px 0">'
            . '<a href="' . esc_url($download_url) . '" style="display:inline-block;background:' . esc_attr($primary) . ';color:#ffffff;text-decoration:none;font-weight:700;font-size:15px;line-height:1;padding:16px 26px;border-radius:14px">' . esc_html($button_label) . '</a>'
            . '</div>'
            . '<p style="margin:18px 0 0 0;font-size:13px;line-height:1.6;color:#8d8e8e">Si al pulsar el botón no se abre directamente en PDF, se mostrará la versión lista para imprimir y guardar como PDF.</p>'
            . '<p style="margin:20px 0 0 0;font-size:15px;line-height:1.6;color:' . esc_attr($primary) . '">Un saludo,<br><strong>' . esc_html($settings['issuer_name']) . '</strong></p>'
            . '</div></div></div></body></html>';
    }

    public static function page_send_document() {
        if (!current_user_can('manage_options')) { wp_die('No autorizado'); }
        $invoice_id = (int) ($_GET['invoice_id'] ?? 0);
        $invoice = self::invoice($invoice_id);
        if (!$invoice) { wp_die('Documento no encontrado'); }
        $client = self::client($invoice['client_id']);
        $contacts = self::client_contacts($client);
        if (!$contacts && !empty($client['email'])) {
            $contacts = [[
                'email' => sanitize_email($client['email']),
                'name' => sanitize_text_field($client['contact_name'] ?? ($client['name'] ?? '')),
                'label' => sanitize_email($client['email']),
            ]];
        }
        echo '<div class="wrap vvd-wrap"><h1>Enviar ' . esc_html(strtolower(self::doc_label($invoice['invoice_type']))) . '</h1><div class="vvd-card">';
        echo '<p class="vvd-note">Elige a qué contacto del cliente quieres enviar este documento.</p>';
        echo '<form method="post" action="' . esc_url(admin_url('admin-post.php')) . '">';
        wp_nonce_field('vvd_send_email_' . $invoice_id);
        echo '<input type="hidden" name="action" value="vvd_send_invoice_email"><input type="hidden" name="invoice_id" value="' . esc_attr($invoice_id) . '">';
        echo '<div class="vvd-grid">';
        echo '<div class="vvd-col-6"><label class="vvd-label">Cliente</label><input class="vvd-input" type="text" value="' . esc_attr($client['name'] ?? '') . '" readonly></div>';
        echo '<div class="vvd-col-6"><label class="vvd-label">Documento</label><input class="vvd-input" type="text" value="' . esc_attr($invoice['code'] ?: self::invoice_code($invoice['series'], $invoice['number'], (int) (self::settings()['series_padding'] ?: 6))) . '" readonly></div>';
        echo '<div class="vvd-col-12"><label class="vvd-label">Enviar a</label><select class="vvd-select" name="recipient_email">';
        foreach ($contacts as $contact) {
            echo '<option value="' . esc_attr($contact['email']) . '">' . esc_html($contact['label']) . '</option>';
        }
        echo '</select></div>';
        echo '<div class="vvd-col-12"><label class="vvd-label">Nombre del destinatario (opcional)</label><input class="vvd-input" type="text" name="recipient_name" value="' . esc_attr($contacts[0]['name'] ?? ($client['contact_name'] ?? '')) . '"></div>';
        echo '</div><p><button class="vvd-btn" type="submit">Enviar correo</button> <a class="vvd-btn light" href="' . esc_url(admin_url('admin.php?page=vvd_facturas&view=' . $invoice_id)) . '">Cancelar</a></p></form></div></div>';
    }

    public static function prepare_invoice_email() {
        $invoice_id = (int) ($_REQUEST['invoice_id'] ?? 0);
        wp_safe_redirect(admin_url('admin.php?page=vvd_enviar_documento&invoice_id=' . $invoice_id));
        exit;
    }

    public static function send_invoice_email() {
        if (!current_user_can('manage_options')) { wp_die('No autorizado'); }
        $invoice_id = (int) ($_POST['invoice_id'] ?? 0);
        check_admin_referer('vvd_send_email_' . $invoice_id);
        global $wpdb;
        $invoice = self::invoice($invoice_id);
        $client = self::client($invoice['client_id']);
        $settings = self::settings();
        $recipient_email = sanitize_email($_POST['recipient_email'] ?? ($_GET['recipient_email'] ?? ($client['email'] ?? '')));
        $recipient_name = sanitize_text_field($_POST['recipient_name'] ?? ($_GET['recipient_name'] ?? ($client['contact_name'] ?? ($client['name'] ?? ''))));
        if (empty($recipient_email)) {
            wp_die('El cliente no tiene ningún email disponible.');
        }

        $code = $invoice['code'] ?: self::invoice_code($invoice['series'], $invoice['number'], (int) ($settings['series_padding'] ?: 6));
        $subject = self::doc_label($invoice['invoice_type']) . ' ' . $code;
        $print_url = self::public_document_url($invoice_id, $invoice);

        $body = self::email_html($invoice, $client, $settings, $print_url);

        $headers = ['Content-Type: text/html; charset=UTF-8'];
        if (!empty($settings['issuer_email'])) {
            $from_name = !empty($settings['issuer_name']) ? wp_specialchars_decode($settings['issuer_name'], ENT_QUOTES) : get_bloginfo('name');
            $headers[] = 'From: ' . $from_name . ' <' . sanitize_email($settings['issuer_email']) . '>';
            $headers[] = 'Reply-To: ' . $from_name . ' <' . sanitize_email($settings['issuer_email']) . '>';
        }

        $attachments = [];
        $pdf = self::maybe_generate_pdf($invoice_id, true);
        if (!empty($pdf['path']) && file_exists($pdf['path']) && (($pdf['engine'] ?? '') === 'html_print')) {
            $attachments[] = $pdf['path'];
        }
        if (($pdf['engine'] ?? '') !== 'html_print') {
            $body .= '<p style="margin:18px 0 0 0;font-size:14px;line-height:1.6;color:#5f6c77"><strong>Nota:</strong> en este servidor no ha sido posible adjuntar automáticamente el PDF visual, pero el botón del correo sigue abriendo la versión correcta lista para imprimir.</p>';
        }

        $sent = wp_mail($recipient_email, $subject, $body, $headers, $attachments);

        if (!$sent && !empty($attachments)) {
            $sent = wp_mail($recipient_email, $subject, $body, $headers);
            if ($sent) {
                self::log_event('email_sent_no_attachment', self::doc_label($invoice['invoice_type']) . ' enviado por email sin adjunto a ' . $recipient_email, $invoice_id, $invoice['client_id'], ['fallback' => 'without_attachment']);
            }
        }

        if ($sent) {
            $wpdb->update(
                self::table('vvd_invoices'),
                [
                    'sent_at' => current_time('mysql'),
                    'status' => ((!empty($invoice['paid_at']) || (($invoice['status'] ?? '') === 'paid')) ? 'paid' : 'sent'),
                    'updated_at' => current_time('mysql'),
                ],
                ['id' => $invoice_id]
            );
            self::log_event('email_sent', self::doc_label($invoice['invoice_type']) . ' enviado por email a ' . $recipient_email, $invoice_id, $invoice['client_id'], ['recipient_name' => $recipient_name]);
        } else {
            self::log_event('email_error', 'Error enviando por email a ' . $recipient_email, $invoice_id, $invoice['client_id'], ['print_url' => $print_url, 'recipient_name' => $recipient_name]);
        }
        wp_safe_redirect(admin_url('admin.php?page=vvd_facturas&view=' . $invoice_id));
        exit;
    }

    public static function toggle_invoice_paid() {
        if (!current_user_can('manage_options')) { wp_die('No autorizado'); }
        $invoice_id = (int) ($_POST['invoice_id'] ?? 0);
        check_admin_referer('vvd_toggle_paid_' . $invoice_id);
        global $wpdb;
        $invoice = self::invoice($invoice_id);
        if (!$invoice || $invoice['invoice_type'] === 'quote') {
            wp_die('Documento no válido.');
        }

        $is_paid = !empty($invoice['paid_at']) || (($invoice['status'] ?? '') === 'paid');
        if ($is_paid) {
            $new_status = !empty($invoice['sent_at']) ? 'sent' : 'issued';
            $wpdb->update(
                self::table('vvd_invoices'),
                [
                    'paid_at' => null,
                    'status' => $new_status,
                    'updated_at' => current_time('mysql'),
                ],
                ['id' => $invoice_id]
            );
            self::log_event('invoice_unpaid', 'Factura marcada como no pagada.', $invoice_id, $invoice['client_id']);
        } else {
            $wpdb->update(
                self::table('vvd_invoices'),
                [
                    'paid_at' => current_time('mysql'),
                    'status' => 'paid',
                    'updated_at' => current_time('mysql'),
                ],
                ['id' => $invoice_id]
            );
            self::log_event('invoice_paid', 'Factura marcada como pagada.', $invoice_id, $invoice['client_id']);
        }

        $redirect = admin_url('admin.php?page=vvd_facturas');
        $args = ['page' => 'vvd_facturas'];
        if (isset($_REQUEST['client_id'])) {
            $args['client_id'] = (int) $_REQUEST['client_id'];
        }
        if (isset($_REQUEST['paid_filter'])) {
            $args['paid_filter'] = sanitize_text_field($_REQUEST['paid_filter']);
        }
        wp_safe_redirect(add_query_arg($args, admin_url('admin.php')));
        exit;
    }

    public static function convert_to_invoice() {
        if (!current_user_can('manage_options')) { wp_die('No autorizado'); }
        $invoice_id = (int) ($_POST['invoice_id'] ?? 0);
        check_admin_referer('vvd_convert_' . $invoice_id);
        global $wpdb;
        $source = self::invoice($invoice_id);
        if (!$source || $source['invoice_type'] !== 'quote') {
            wp_die('Solo se puede convertir un presupuesto.');
        }
        $lines = self::invoice_lines($invoice_id);
        $series_data = self::reserve_series_number('standard');
        $settings = self::settings();
        $tax_rate = (float) ($settings['default_tax_rate'] ?: 21);
        $irpf_rate = (float) ($settings['default_irpf_rate'] ?: 15);
        $base = (float) ($source['taxable_base'] ?? 0);
        $tax_amount = round($base * $tax_rate / 100, 2);
        $irpf_amount = round($base * $irpf_rate / 100, 2);
        $total_amount = round($base + $tax_amount - $irpf_amount, 2);
        $hash_prev = self::latest_hash_before();
        $hash_current = self::compute_invoice_hash([
            'series' => $series_data['series'], 'number' => $series_data['number'], 'issue_date' => current_time('Y-m-d'), 'client_id' => $source['client_id'],
            'taxable_base' => $base, 'tax_amount' => $tax_amount, 'irpf_amount' => $irpf_amount, 'total_amount' => $total_amount,
        ], $lines, $hash_prev);
        $wpdb->insert(self::table('vvd_invoices'), [
            'series' => $series_data['series'],
            'series_label' => $series_data['series_label'],
            'number' => $series_data['number'],
            'code' => $series_data['code'],
            'invoice_type' => 'standard',
            'issue_date' => current_time('Y-m-d'),
            'client_id' => $source['client_id'],
            'original_invoice_id' => $source['id'],
            'notes' => trim("Convertida desde presupuesto " . ($source['code'] ?: $invoice_id) . "
" . $source['notes']),
            'taxable_base' => $base,
            'tax_rate' => $tax_rate,
            'tax_amount' => $tax_amount,
            'irpf_rate' => $irpf_rate,
            'irpf_amount' => $irpf_amount,
            'total_amount' => $total_amount,
            'status' => 'issued',
            'verifactu_status' => 'draft-ready',
            'hash_prev' => $hash_prev,
            'hash_current' => $hash_current,
            'created_at' => current_time('mysql'),
            'updated_at' => current_time('mysql'),
        ]);
        $new_id = (int) $wpdb->insert_id;
        foreach ($lines as $line) {
            unset($line['id']);
            $wpdb->insert(self::table('vvd_invoice_lines'), ['invoice_id' => $new_id] + $line);
        }
        self::log_event('quote_converted', 'Presupuesto convertido en factura: ' . $series_data['code'], $new_id, $source['client_id'], ['source_quote_id' => $invoice_id]);
        self::maybe_generate_pdf($new_id, true);
        wp_safe_redirect(admin_url('admin.php?page=vvd_facturas&view=' . $new_id));
        exit;
    }

    public static function delete_invoice() {
        if (!current_user_can('manage_options')) { wp_die('No autorizado'); }
        $invoice_id = (int) ($_POST['invoice_id'] ?? 0);
        check_admin_referer('vvd_delete_invoice_' . $invoice_id);
        global $wpdb;
        $invoice = self::invoice($invoice_id);
        if (!$invoice) {
            wp_safe_redirect(admin_url('admin.php?page=vvd_facturas'));
            exit;
        }
        $wpdb->delete(self::table('vvd_invoice_lines'), ['invoice_id' => $invoice_id]);
        $wpdb->delete(self::table('vvd_invoices'), ['id' => $invoice_id]);
        self::log_event('invoice_deleted', 'Documento eliminado: ' . ($invoice['code'] ?: $invoice_id), $invoice_id, $invoice['client_id']);
        wp_safe_redirect(admin_url('admin.php?page=vvd_facturas'));
        exit;
    }


    protected static function accounting_entries_table() {
        return self::table('vvd_accounting_entries');
    }

    protected static function imports_table() {
        return self::table('vvd_imports');
    }

    protected static function month_key_from_date($date) {
        $ts = strtotime((string) $date);
        if (!$ts) {
            return '';
        }
        return gmdate('Y-m', $ts);
    }

    protected static function month_diff($ymA, $ymB) {
        if (!$ymA || !$ymB) {
            return 999;
        }
        [$y1, $m1] = array_map('intval', explode('-', $ymA));
        [$y2, $m2] = array_map('intval', explode('-', $ymB));
        return (($y2 - $y1) * 12) + ($m2 - $m1);
    }

    protected static function recurring_description_key($description) {
        $value = strtoupper(remove_accents((string) $description));
        $value = preg_replace('/(FAC|RT|KD)[- ]?[0-9A-Z-]+/', ' ', $value);
        $value = preg_replace('/\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4}/', ' ', $value);
        $value = preg_replace('/\d{4,}/', ' ', $value);
        $value = preg_replace('/[^A-Z ]+/', ' ', $value);
        $value = preg_replace('/\s+/', ' ', trim($value));
        return $value !== '' ? $value : 'SIN DESCRIPCION';
    }

    protected static function recurring_group_key($entry) {
        $nif = strtoupper(trim((string) ($entry['nif'] ?? '')));
        $account = preg_replace('/\s+/', '', (string) ($entry['account_code'] ?? ''));
        $account_prefix = substr($account, 0, 3);
        $desc = self::recurring_description_key($entry['description'] ?? '');
        return implode('|', [
            $nif !== '' ? $nif : 'NO-NIF',
            $account_prefix !== '' ? $account_prefix : '000',
            $desc,
        ]);
    }

    protected static function build_empty_monthly_summary($reference_ym, $next_ym) {
        return [
            'current_expected_total' => 0.0,
            'current_recorded_total' => 0.0,
            'current_pending_total' => 0.0,
            'next_expected_total' => 0.0,
            'count' => 0,
            'current_month' => $reference_ym,
            'next_month' => $next_ym,
            'items' => [],
        ];
    }

    protected static function build_empty_annual_summary($reference_year, $reference_month_num = 1) {
        $reference_month_num = max(1, min(12, (int) $reference_month_num));
        $current_month_label = date_i18n('F', strtotime(sprintf('%04d-%02d-01', max(2000, (int) $reference_year), $reference_month_num)));
        return [
            'current_expected_total' => 0.0,
            'current_recorded_total' => 0.0,
            'current_pending_total' => 0.0,
            'next_expected_total' => 0.0,
            'count' => 0,
            'current_year' => (int) $reference_year,
            'next_year' => (int) $reference_year + 1,
            'reference_month' => $reference_month_num,
            'reference_month_label' => $current_month_label,
            'items' => [],
        ];
    }

    protected static function detect_recurring_expenses($months_back = 18, $reference_ym = null) {
        global $wpdb;
        $entries_table = self::accounting_entries_table();
        $base_ts = current_time('timestamp');
        if (is_string($reference_ym) && preg_match('/^\d{4}-\d{2}$/', $reference_ym)) {
            $candidate_ts = strtotime($reference_ym . '-01');
            if ($candidate_ts) {
                $base_ts = $candidate_ts;
            }
        }

        $reference_ym = gmdate('Y-m', $base_ts);
        $reference_year = (int) gmdate('Y', $base_ts);
        $reference_month_num = (int) gmdate('n', $base_ts);
        $next_ym = gmdate('Y-m', strtotime('+1 month', strtotime($reference_ym . '-01')));
        $start_date = gmdate('Y-m-01', strtotime('-' . max(1, (int) $months_back) . ' months', $base_ts));
        $rows = $wpdb->get_results($wpdb->prepare(
            "SELECT entry_date, nif, account_code, description, total_amount, invoice_number_raw FROM {$entries_table} WHERE entry_group='expense' AND entry_date >= %s ORDER BY entry_date ASC, id ASC",
            $start_date
        ), ARRAY_A);

        $groups = [];
        foreach ((array) $rows as $row) {
            $key = self::recurring_group_key($row);
            $ym = self::month_key_from_date($row['entry_date'] ?? '');
            if ($ym === '') {
                continue;
            }
            $amount = abs((float) ($row['total_amount'] ?? 0));
            if ($amount <= 0) {
                continue;
            }
            $row_ts = strtotime((string) ($row['entry_date'] ?? ''));
            if (!$row_ts) {
                continue;
            }
            $day = (int) gmdate('j', $row_ts);
            $month_num = (int) gmdate('n', $row_ts);
            $year_num = (int) gmdate('Y', $row_ts);
            if (!isset($groups[$key])) {
                $groups[$key] = [
                    'key' => $key,
                    'label' => trim((string) ($row['description'] ?? '')) ?: 'Gasto recurrente',
                    'nif' => trim((string) ($row['nif'] ?? '')),
                    'account_code' => trim((string) ($row['account_code'] ?? '')),
                    'months' => [],
                    'days' => [],
                    'month_totals' => [],
                    'month_days' => [],
                    'annual_months' => [],
                    'annual_totals' => [],
                    'annual_days' => [],
                ];
            }
            $groups[$key]['months'][$ym] = true;
            $groups[$key]['days'][] = $day;
            if (!isset($groups[$key]['month_totals'][$ym])) {
                $groups[$key]['month_totals'][$ym] = 0.0;
            }
            if (!isset($groups[$key]['month_days'][$ym])) {
                $groups[$key]['month_days'][$ym] = [];
            }
            $groups[$key]['month_totals'][$ym] += $amount;
            $groups[$key]['month_days'][$ym][] = $day;

            if (!isset($groups[$key]['annual_months'][$month_num])) {
                $groups[$key]['annual_months'][$month_num] = [];
            }
            $groups[$key]['annual_months'][$month_num][$year_num] = true;
            if (!isset($groups[$key]['annual_totals'][$month_num])) {
                $groups[$key]['annual_totals'][$month_num] = [];
            }
            if (!isset($groups[$key]['annual_totals'][$month_num][$year_num])) {
                $groups[$key]['annual_totals'][$month_num][$year_num] = 0.0;
            }
            if (!isset($groups[$key]['annual_days'][$month_num])) {
                $groups[$key]['annual_days'][$month_num] = [];
            }
            $groups[$key]['annual_totals'][$month_num][$year_num] += $amount;
            $groups[$key]['annual_days'][$month_num][] = $day;
        }

        $monthly_detected = [];
        $annual_detected = [];

        foreach ($groups as $group) {
            $months = array_keys((array) $group['months']);
            sort($months);
            $distinct_months = count($months);

            if ($distinct_months >= 2) {
                $gaps = [];
                for ($i = 1; $i < $distinct_months; $i++) {
                    $gaps[] = self::month_diff($months[$i - 1], $months[$i]);
                }
                $avg_gap = $gaps ? (array_sum($gaps) / count($gaps)) : 99;

                $month_totals = array_values(array_filter(array_map('floatval', (array) $group['month_totals']), function($v){ return $v > 0; }));
                if ($month_totals) {
                    sort($month_totals);
                    $avg_month_amount = array_sum($month_totals) / count($month_totals);
                    $median_month_amount = $month_totals[(int) floor((count($month_totals) - 1) / 2)];
                    if ((count($month_totals) % 2) === 0 && count($month_totals) > 1) {
                        $mid = (int) (count($month_totals) / 2);
                        $median_month_amount = ($month_totals[$mid - 1] + $month_totals[$mid]) / 2;
                    }
                    $estimated_amount = $median_month_amount > 0 ? $median_month_amount : $avg_month_amount;
                    $max_amount = max($month_totals);
                    $min_amount = min($month_totals);
                    $variation_ratio = $estimated_amount > 0 ? (($max_amount - $min_amount) / $estimated_amount) : 999;

                    $is_monthly = ($distinct_months >= 3 && $avg_gap <= 1.6) || ($distinct_months === 2 && (int) ($gaps[0] ?? 99) === 1);
                    if ($is_monthly && $variation_ratio <= 0.90) {
                        $days = array_map('intval', (array) $group['days']);
                        sort($days);
                        $expected_day = $days ? (int) round(array_sum($days) / count($days)) : 1;
                        $expected_day = max(1, min(28, $expected_day));

                        $reference_recorded = (float) ($group['month_totals'][$reference_ym] ?? 0);
                        $next_recorded = (float) ($group['month_totals'][$next_ym] ?? 0);
                        $last_seen = end($months) ?: '';
                        $first_seen = reset($months) ?: '';
                        $gap_to_reference = self::month_diff($last_seen, $reference_ym);
                        $gap_from_first = self::month_diff($first_seen, $reference_ym);
                        $should_show_reference = ($gap_from_first >= 0 && $gap_to_reference >= 0 && $gap_to_reference <= 2);
                        $should_show_next = ($gap_from_first >= -1 && self::month_diff($last_seen, $next_ym) >= 0 && self::month_diff($last_seen, $next_ym) <= 2);

                        if ($should_show_reference || $should_show_next) {
                            $reference_expected = $should_show_reference ? $estimated_amount : 0;
                            $pending_reference = max(0, $reference_expected - $reference_recorded);
                            $next_expected = $should_show_next ? max($estimated_amount, $next_recorded) : 0;
                            $monthly_detected[] = [
                                'label' => $group['label'],
                                'nif' => $group['nif'],
                                'account_code' => $group['account_code'],
                                'months_count' => $distinct_months,
                                'expected_day' => $expected_day,
                                'estimated_amount' => round($estimated_amount, 2),
                                'average_amount' => round($avg_month_amount, 2),
                                'current_recorded' => round($reference_recorded, 2),
                                'current_expected' => round($reference_expected, 2),
                                'current_pending' => round($pending_reference, 2),
                                'next_expected' => round($next_expected, 2),
                                'last_seen_month' => $last_seen,
                                'first_seen_month' => $first_seen,
                            ];
                        }
                    }
                }
            }

            foreach ((array) $group['annual_months'] as $month_num => $years_map) {
                $years = array_keys((array) $years_map);
                sort($years);
                $years_count = count($years);
                if ($years_count < 2) {
                    continue;
                }
                $totals_by_year = array_values(array_filter(array_map('floatval', (array) ($group['annual_totals'][$month_num] ?? [])), function($v){ return $v > 0; }));
                if (!$totals_by_year) {
                    continue;
                }
                sort($totals_by_year);
                $avg_amount = array_sum($totals_by_year) / count($totals_by_year);
                $median_amount = $totals_by_year[(int) floor((count($totals_by_year) - 1) / 2)];
                if ((count($totals_by_year) % 2) === 0 && count($totals_by_year) > 1) {
                    $mid = (int) (count($totals_by_year) / 2);
                    $median_amount = ($totals_by_year[$mid - 1] + $totals_by_year[$mid]) / 2;
                }
                $estimated_amount = $median_amount > 0 ? $median_amount : $avg_amount;
                $max_amount = max($totals_by_year);
                $min_amount = min($totals_by_year);
                $variation_ratio = $estimated_amount > 0 ? (($max_amount - $min_amount) / $estimated_amount) : 999;
                if ($variation_ratio > 1.20) {
                    continue;
                }

                $days = array_map('intval', (array) ($group['annual_days'][$month_num] ?? []));
                sort($days);
                $expected_day = $days ? (int) round(array_sum($days) / count($days)) : 1;
                $expected_day = max(1, min(28, $expected_day));
                $last_seen_year = (int) end($years);
                $last_seen_month = sprintf('%04d-%02d', $last_seen_year, (int) $month_num);
                if ((int) $month_num !== $reference_month_num) {
                    continue;
                }
                $current_recorded = (float) ($group['annual_totals'][$month_num][$reference_year] ?? 0);
                $next_recorded = (float) ($group['annual_totals'][$month_num][$reference_year + 1] ?? 0);
                $should_show_current = ($reference_year - $last_seen_year <= 1);
                $reference_expected = $should_show_current ? $estimated_amount : 0;
                $current_pending = max(0, $reference_expected - $current_recorded);
                $next_expected = max($estimated_amount, $next_recorded);
                $annual_month_label = date_i18n('F', strtotime(sprintf('%04d-%02d-01', max(2000, $reference_year), (int) $month_num)));

                $annual_detected[] = [
                    'label' => $group['label'],
                    'nif' => $group['nif'],
                    'account_code' => $group['account_code'],
                    'years_count' => $years_count,
                    'annual_month' => (int) $month_num,
                    'annual_month_label' => $annual_month_label,
                    'expected_day' => $expected_day,
                    'estimated_amount' => round($estimated_amount, 2),
                    'average_amount' => round($avg_amount, 2),
                    'current_recorded' => round($current_recorded, 2),
                    'current_expected' => round($reference_expected, 2),
                    'current_pending' => round($current_pending, 2),
                    'next_expected' => round($next_expected, 2),
                    'last_seen_month' => $last_seen_month,
                ];
            }
        }

        usort($monthly_detected, function($a, $b){
            return ($b['estimated_amount'] <=> $a['estimated_amount']);
        });
        usort($annual_detected, function($a, $b){
            if ((int) $a['annual_month'] === (int) $b['annual_month']) {
                return ($b['estimated_amount'] <=> $a['estimated_amount']);
            }
            return ((int) $a['annual_month'] <=> (int) $b['annual_month']);
        });

        $monthly_summary = self::build_empty_monthly_summary($reference_ym, $next_ym);
        $monthly_summary['count'] = count($monthly_detected);
        $monthly_summary['items'] = $monthly_detected;
        foreach ($monthly_detected as $item) {
            $monthly_summary['current_expected_total'] += (float) $item['current_expected'];
            $monthly_summary['current_recorded_total'] += (float) $item['current_recorded'];
            $monthly_summary['current_pending_total'] += (float) $item['current_pending'];
            $monthly_summary['next_expected_total'] += (float) $item['next_expected'];
        }

        $annual_summary = self::build_empty_annual_summary($reference_year, $reference_month_num);
        $annual_summary['count'] = count($annual_detected);
        $annual_summary['items'] = $annual_detected;
        foreach ($annual_detected as $item) {
            $annual_summary['current_expected_total'] += (float) $item['current_expected'];
            $annual_summary['current_recorded_total'] += (float) $item['current_recorded'];
            $annual_summary['current_pending_total'] += (float) $item['current_pending'];
            $annual_summary['next_expected_total'] += (float) $item['next_expected'];
        }

        return [
            'monthly' => $monthly_summary,
            'annual' => $annual_summary,
        ];
    }

    protected static function normalize_amount($value) {
        if (is_int($value) || is_float($value)) {
            return (float) $value;
        }
        $value = trim((string) $value);
        if ($value === '') {
            return 0.0;
        }
        $value = str_replace(["\xc2\xa0", '€', 'EUR', ' '], '', $value);
        if (strpos($value, ',') !== false) {
            $value = str_replace('.', '', $value);
            $value = str_replace(',', '.', $value);
        }
        $value = preg_replace('/[^0-9.\-]/', '', $value);
        return (float) $value;
    }

    protected static function normalize_csv_date($value) {
        $value = trim((string) $value);
        if ($value === '') {
            return current_time('Y-m-d');
        }
        $parts = preg_split('/[\/\-]/', $value);
        if (count($parts) === 3) {
            if (strlen($parts[2]) === 4) {
                return sprintf('%04d-%02d-%02d', (int) $parts[2], (int) $parts[1], (int) $parts[0]);
            }
            return sprintf('%04d-%02d-%02d', (int) $parts[0], (int) $parts[1], (int) $parts[2]);
        }
        return current_time('Y-m-d');
    }

    protected static function normalize_invoice_number($value) {
        $value = strtoupper(trim((string) $value));
        $value = preg_replace('/\s+/', '', $value);
        return $value;
    }

    protected static function classify_accounting_row($row, $filename = '') {
        $account = preg_replace('/\s+/', '', (string) ($row['Cuenta'] ?? ''));
        $number_raw = trim((string) ($row['Nº factura'] ?? ''));
        $number = self::normalize_invoice_number($number_raw);
        $filename_upper = strtoupper((string) $filename);

        $entry_group = 'other';
        if (strpos($account, '430') === 0) {
            $entry_group = 'income';
        } elseif (strpos($account, '410') === 0) {
            $entry_group = 'expense';
        }

        $document_type = $entry_group === 'expense' ? 'expense_document' : 'invoice';
        $business_line = 'other';
        $series_code = 'NONE';

        if ($number !== '') {
            if (strpos($number, 'RT-') === 0) {
                $document_type = 'rectificative';
                $series_code = 'RT';
            } elseif (strpos($number, 'FAC-') === 0) {
                $document_type = 'invoice';
                $business_line = 'agency';
                $series_code = 'FAC';
            } elseif (strpos($number, 'KD') === 0) {
                $document_type = 'invoice';
                $business_line = 'kit_digital';
                $series_code = 'KD';
            } elseif (preg_match('/^\d{3}$/', $number)) {
                $document_type = 'invoice';
                $business_line = 'kit_digital';
                $series_code = 'NUMERIC_3';
            } elseif (preg_match('/^[A-Z]{2,5}\-/', $number, $m)) {
                $series_code = strtoupper(rtrim($m[0], '-'));
            }
        }

        if ($entry_group === 'expense') {
            $document_type = 'expense_document';
            $business_line = 'expense';
            if ($series_code === 'NONE' && $number !== '') {
                $series_code = 'SUPPLIER';
            }
        }

        $source_period_type = 'manual_import';
        if ($entry_group === 'expense') {
            $source_period_type = 'quarterly_expense_import';
        } elseif (strpos($filename_upper, '(1)') === false) {
            $source_period_type = 'historical_income';
        }

        return [
            'entry_group' => $entry_group,
            'document_type' => $document_type,
            'business_line' => $business_line,
            'series_code' => $series_code,
            'source_period_type' => $source_period_type,
        ];
    }

    protected static function parse_accounting_csv($path, $filename) {
        $rows = [];
        $errors = [];
        if (!file_exists($path)) {
            return ['rows' => [], 'errors' => ['Archivo no encontrado.']];
        }
        if (($handle = fopen($path, 'r')) === false) {
            return ['rows' => [], 'errors' => ['No se ha podido abrir el CSV.']];
        }
        $headers = fgetcsv($handle, 0, ',', '"', '\\');
        if (!$headers) {
            fclose($handle);
            return ['rows' => [], 'errors' => ['CSV vacío o sin cabecera.']];
        }
        $headers = array_map(function($h){ return trim((string) $h); }, $headers);
        $line = 1;
        while (($data = fgetcsv($handle, 0, ',', '"', '\\')) !== false) {
            $line++;
            if ($data === [null] || $data === false) {
                continue;
            }
            $assoc = [];
            foreach ($headers as $idx => $header) {
                $assoc[$header] = isset($data[$idx]) ? trim((string) $data[$idx]) : '';
            }
            $desc = strtoupper(trim((string) ($assoc['Descripción'] ?? '')));
            if ($desc === 'SUBTOTAL' || $desc === '') {
                continue;
            }
            $classified = self::classify_accounting_row($assoc, $filename);
            $rows[] = [
                'entry_date' => self::normalize_csv_date($assoc['Fecha'] ?? ''),
                'nif' => sanitize_text_field($assoc['NIF'] ?? ''),
                'account_code' => sanitize_text_field($assoc['Cuenta'] ?? ''),
                'description' => sanitize_text_field($assoc['Descripción'] ?? ''),
                'journal_entry' => sanitize_text_field($assoc['Asiento'] ?? ''),
                'invoice_number_raw' => sanitize_text_field($assoc['Nº factura'] ?? ''),
                'invoice_number_normalized' => sanitize_text_field(self::normalize_invoice_number($assoc['Nº factura'] ?? '')),
                'base_amount' => self::normalize_amount($assoc['Base'] ?? 0),
                'tax_rate' => self::normalize_amount($assoc['Tipo (%)'] ?? 0),
                'tax_amount' => self::normalize_amount($assoc['Cuota'] ?? 0),
                'irpf_rate' => self::normalize_amount($assoc['IRPF (%)'] ?? 0),
                'irpf_amount' => self::normalize_amount($assoc['Total IRPF'] ?? 0),
                'total_amount' => self::normalize_amount($assoc['Total'] ?? 0),
                'entry_group' => $classified['entry_group'],
                'document_type' => $classified['document_type'],
                'business_line' => $classified['business_line'],
                'series_code' => $classified['series_code'],
                'source_file' => sanitize_file_name($filename),
                'source_period_type' => $classified['source_period_type'],
                'raw_row_json' => wp_json_encode($assoc),
                'created_at' => current_time('mysql'),
            ];
        }
        fclose($handle);
        return ['rows' => $rows, 'errors' => $errors];
    }

    protected static function maybe_sync_client_from_entry($entry) {
        if (($entry['entry_group'] ?? '') !== 'income') {
            return;
        }
        global $wpdb;
        $name = sanitize_text_field($entry['description'] ?? '');
        $tax_id = sanitize_text_field($entry['nif'] ?? '');
        if ($name === '' && $tax_id === '') {
            return;
        }
        $clients_table = self::table('vvd_clients');
        $existing_id = 0;
        if ($tax_id !== '') {
            $existing_id = (int) $wpdb->get_var($wpdb->prepare("SELECT id FROM {$clients_table} WHERE tax_id=%s LIMIT 1", $tax_id));
        }
        if (!$existing_id && $name !== '') {
            $existing_id = (int) $wpdb->get_var($wpdb->prepare("SELECT id FROM {$clients_table} WHERE name=%s LIMIT 1", $name));
        }
        if ($existing_id) {
            $wpdb->update($clients_table, [
                'name' => $name ?: $wpdb->get_var($wpdb->prepare("SELECT name FROM {$clients_table} WHERE id=%d", $existing_id)),
                'tax_id' => $tax_id ?: $wpdb->get_var($wpdb->prepare("SELECT tax_id FROM {$clients_table} WHERE id=%d", $existing_id)),
                'updated_at' => current_time('mysql'),
            ], ['id' => $existing_id]);
        } else {
            $wpdb->insert($clients_table, [
                'name' => $name ?: 'Cliente importado',
                'tax_id' => $tax_id ?: 'PENDIENTE',
                'created_at' => current_time('mysql'),
                'updated_at' => current_time('mysql'),
            ]);
        }
    }


    protected static function analysis_income_imported_monthly($year) {
        global $wpdb;
        $entries_table = self::accounting_entries_table();
        $invoices_table = self::table('vvd_invoices');
        $rows = $wpdb->get_results($wpdb->prepare(
            "SELECT MONTH(e.entry_date) as mm, SUM(e.base_amount) as total FROM {$entries_table} e LEFT JOIN {$invoices_table} i ON UPPER(TRIM(i.code)) = UPPER(TRIM(e.invoice_number_normalized)) WHERE e.entry_group='income' AND YEAR(e.entry_date) = %d AND i.id IS NULL GROUP BY MONTH(e.entry_date) ORDER BY mm ASC",
            $year
        ), ARRAY_A);
        $bucket = [];
        foreach ((array) $rows as $row) {
            $month_num = (int) ($row['mm'] ?? 0);
            if ($month_num >= 1 && $month_num <= 12) {
                $bucket[$month_num] = (float) ($row['total'] ?? 0);
            }
        }
        return $bucket;
    }

    protected static function analysis_income_generated_monthly($year) {
        global $wpdb;
        $invoices_table = self::table('vvd_invoices');
        $rows = $wpdb->get_results($wpdb->prepare(
            "SELECT MONTH(issue_date) as mm, SUM(taxable_base) as total FROM {$invoices_table} WHERE invoice_type != 'quote' AND YEAR(issue_date) = %d GROUP BY MONTH(issue_date) ORDER BY mm ASC",
            $year
        ), ARRAY_A);
        $bucket = [];
        foreach ((array) $rows as $row) {
            $month_num = (int) ($row['mm'] ?? 0);
            if ($month_num >= 1 && $month_num <= 12) {
                $bucket[$month_num] = (float) ($row['total'] ?? 0);
            }
        }
        return $bucket;
    }

    protected static function analysis_expense_monthly($year) {
        global $wpdb;
        $entries_table = self::accounting_entries_table();
        $rows = $wpdb->get_results($wpdb->prepare(
            "SELECT MONTH(entry_date) as mm, SUM(base_amount) as total FROM {$entries_table} WHERE entry_group='expense' AND YEAR(entry_date) = %d GROUP BY MONTH(entry_date) ORDER BY mm ASC",
            $year
        ), ARRAY_A);
        $bucket = [];
        foreach ((array) $rows as $row) {
            $month_num = (int) ($row['mm'] ?? 0);
            if ($month_num >= 1 && $month_num <= 12) {
                $bucket[$month_num] = (float) ($row['total'] ?? 0);
            }
        }
        return $bucket;
    }

    protected static function analysis_monthly_series($year = null) {
        $year = (int) ($year ?: current_time('Y'));
        if ($year < 2000) {
            $year = (int) current_time('Y');
        }

        $imported_income = self::analysis_income_imported_monthly($year);
        $generated_income = self::analysis_income_generated_monthly($year);
        $expenses = self::analysis_expense_monthly($year);

        $series = [];
        for ($month_num = 1; $month_num <= 12; $month_num++) {
            $ts = strtotime(sprintf('%04d-%02d-01', $year, $month_num));
            $series[] = [
                'ym' => sprintf('%04d-%02d', $year, $month_num),
                'label' => date_i18n('M y', $ts),
                'income' => (float) ($imported_income[$month_num] ?? 0) + (float) ($generated_income[$month_num] ?? 0),
                'expense' => (float) ($expenses[$month_num] ?? 0),
            ];
        }
        return $series;
    }

    protected static function analysis_yearly_history($years = 6) {
        global $wpdb;
        $entries_table = self::accounting_entries_table();
        $invoices_table = self::table('vvd_invoices');
        $from_year = max(2000, (int) current_time('Y') - max(2, (int) $years));

        $imported_rows = $wpdb->get_results($wpdb->prepare(
            "SELECT YEAR(e.entry_date) as yy, SUM(e.base_amount) as income_total FROM {$entries_table} e LEFT JOIN {$invoices_table} i ON UPPER(TRIM(i.code)) = UPPER(TRIM(e.invoice_number_normalized)) WHERE e.entry_group='income' AND YEAR(e.entry_date) >= %d AND i.id IS NULL GROUP BY YEAR(e.entry_date) ORDER BY yy ASC",
            $from_year
        ), ARRAY_A);
        $expense_rows = $wpdb->get_results($wpdb->prepare(
            "SELECT YEAR(entry_date) as yy, SUM(base_amount) as expense_total FROM {$entries_table} WHERE entry_group='expense' AND YEAR(entry_date) >= %d GROUP BY YEAR(entry_date) ORDER BY yy ASC",
            $from_year
        ), ARRAY_A);
        $generated_rows = $wpdb->get_results($wpdb->prepare(
            "SELECT YEAR(issue_date) as yy, SUM(taxable_base) as income_total FROM {$invoices_table} WHERE invoice_type != 'quote' AND YEAR(issue_date) >= %d GROUP BY YEAR(issue_date) ORDER BY yy ASC",
            $from_year
        ), ARRAY_A);

        $bucket = [];
        foreach ((array) $imported_rows as $row) {
            $yy = (int) ($row['yy'] ?? 0);
            if ($yy > 0) {
                $bucket[$yy]['income'] = (float) ($bucket[$yy]['income'] ?? 0) + (float) ($row['income_total'] ?? 0);
            }
        }
        foreach ((array) $generated_rows as $row) {
            $yy = (int) ($row['yy'] ?? 0);
            if ($yy > 0) {
                $bucket[$yy]['income'] = (float) ($bucket[$yy]['income'] ?? 0) + (float) ($row['income_total'] ?? 0);
            }
        }
        foreach ((array) $expense_rows as $row) {
            $yy = (int) ($row['yy'] ?? 0);
            if ($yy > 0) {
                $bucket[$yy]['expense'] = (float) ($bucket[$yy]['expense'] ?? 0) + (float) ($row['expense_total'] ?? 0);
            }
        }

        ksort($bucket);
        $series = [];
        foreach ($bucket as $yy => $row) {
            $income = (float) ($row['income'] ?? 0);
            $expense = (float) ($row['expense'] ?? 0);
            $series[] = [
                'year' => (int) $yy,
                'income' => $income,
                'expense' => $expense,
                'net' => $income - $expense,
            ];
        }
        return $series;
    }

    protected static function totals_from_monthly_series($series) {
        $totals = ['income' => 0.0, 'expense' => 0.0, 'months_with_income' => 0, 'months_with_expense' => 0];
        foreach ((array) $series as $row) {
            $income = (float) ($row['income'] ?? 0);
            $expense = (float) ($row['expense'] ?? 0);
            $totals['income'] += $income;
            $totals['expense'] += $expense;
            if ($income > 0) { $totals['months_with_income']++; }
            if ($expense > 0) { $totals['months_with_expense']++; }
        }
        return $totals;
    }

    protected static function render_analysis_month_chart($series, $year = null) {
        $max = 0.0;
        $sum_income = 0.0;
        $sum_expense = 0.0;
        foreach ((array) $series as $row) {
            $max = max($max, (float) ($row['income'] ?? 0), (float) ($row['expense'] ?? 0));
            $sum_income += (float) ($row['income'] ?? 0);
            $sum_expense += (float) ($row['expense'] ?? 0);
        }
        $max = max($max, 1);
        ob_start();
        echo '<div class="vvd-chart-card">';
        echo '<div class="vvd-chart-header"><div><h2>Ingresos y gastos</h2><p>Ejercicio ' . esc_html((string) ($year ?: current_time('Y'))) . ' por meses · ingresos = históricas importadas + facturas generadas</p></div><span class="vvd-pill">Visual</span></div>';
        echo '<div class="vvd-chart-legend"><span><i class="income"></i>Ingresos</span><span><i class="expense"></i>Gastos</span></div>';
        echo '<div class="vvd-chart-summary"><div><span>Ingresos ' . esc_html((string) ($year ?: current_time('Y'))) . '</span><strong>' . esc_html(number_format($sum_income, 2, ',', '.')) . ' €</strong></div><div><span>Gastos ' . esc_html((string) ($year ?: current_time('Y'))) . '</span><strong>' . esc_html(number_format($sum_expense, 2, ',', '.')) . ' €</strong></div></div>';
        echo '<div class="vvd-bars-wrap">';
        foreach ((array) $series as $row) {
            $income_h = round((((float) $row['income']) / $max) * 180, 1);
            $expense_h = round((((float) $row['expense']) / $max) * 180, 1);
            echo '<div class="vvd-bar-group">';
            echo '<div class="vvd-bar-stack">';
            echo '<div class="vvd-bar income" title="Ingresos ' . esc_attr(number_format((float) $row['income'], 2, ',', '.')) . ' €" style="height:' . esc_attr($income_h) . 'px"></div>';
            echo '<div class="vvd-bar expense" title="Gastos ' . esc_attr(number_format((float) $row['expense'], 2, ',', '.')) . ' €" style="height:' . esc_attr($expense_h) . 'px"></div>';
            echo '</div>';
            echo '<div class="vvd-bar-label">' . esc_html($row['label']) . '</div>';
            echo '</div>';
        }
        echo '</div>';
        echo '<p class="vvd-chart-foot">Valores sin impuestos. Verde entra, rojo sale. Contablemente no es poesía, pero se entiende rápido.</p>';
        echo '</div>';
        return ob_get_clean();
    }

    protected static function render_analysis_history_chart($series, $current_year_income = 0.0, $current_year = null) {
        $max = 0.0;
        foreach ((array) $series as $row) {
            $max = max($max, (float) ($row['income'] ?? 0));
        }
        $max = max($max, 1);
        ob_start();
        echo '<div class="vvd-chart-card">';
        echo '<div class="vvd-chart-header"><div><h2>Histórico de facturación</h2><p>Base imponible por año, sumando históricas importadas + facturas generadas</p></div><span class="vvd-pill">Histórico</span></div>';
        echo '<div class="vvd-chart-summary"><div><span>Este año (' . esc_html((string) ($current_year ?: current_time('Y'))) . ')</span><strong>' . esc_html(number_format((float) $current_year_income, 2, ',', '.')) . ' €</strong></div><div><span>Referencia</span><strong>Base imponible</strong></div></div>';
        echo '<div class="vvd-history-wrap">';
        foreach ((array) $series as $row) {
            $height = round((((float) $row['income']) / $max) * 200, 1);
            $net_class = ((float) ($row['net'] ?? 0)) >= 0 ? 'positive' : 'negative';
            echo '<div class="vvd-history-col">';
            echo '<div class="vvd-history-bar ' . esc_attr($net_class) . '" title="' . esc_attr(number_format((float) $row['income'], 2, ',', '.')) . ' €" style="height:' . esc_attr($height) . 'px"></div>';
            echo '<div class="vvd-history-year">' . esc_html((string) $row['year']) . '</div>';
            echo '<div class="vvd-history-amount">' . esc_html(number_format((float) $row['income'], 0, ',', '.')) . ' €</div>';
            echo '</div>';
        }
        if (empty($series)) {
            echo '<p>No hay todavía histórico suficiente para dibujar la película.</p>';
        }
        echo '</div>';
        echo '</div>';
        return ob_get_clean();
    }


    protected static function analysis_available_years() {
        global $wpdb;
        $entries_table = self::accounting_entries_table();
        $invoices_table = self::table('vvd_invoices');
        $years = [];

        $entry_years = $wpdb->get_col("SELECT DISTINCT YEAR(entry_date) as yy FROM {$entries_table} WHERE entry_date IS NOT NULL AND entry_date != '' ORDER BY yy DESC");
        $invoice_years = $wpdb->get_col("SELECT DISTINCT YEAR(issue_date) as yy FROM {$invoices_table} WHERE invoice_type != 'quote' AND issue_date IS NOT NULL AND issue_date != '' ORDER BY yy DESC");

        foreach (array_merge((array) $entry_years, (array) $invoice_years) as $yy) {
            $yy = (int) $yy;
            if ($yy >= 2000) {
                $years[$yy] = $yy;
            }
        }

        if (empty($years)) {
            $current = (int) current_time('Y');
            $years[$current] = $current;
        }

        krsort($years);
        return array_values($years);
    }

    protected static function analysis_documents_list($type_filter = 'all', $year_filter = 'all') {
        global $wpdb;
        $entries_table = self::accounting_entries_table();
        $invoices_table = self::table('vvd_invoices');
        $clients_table = self::table('vvd_clients');

        $type_filter = in_array($type_filter, ['all', 'sale', 'purchase'], true) ? $type_filter : 'all';
        $year_filter = ($year_filter === 'all') ? 'all' : (int) $year_filter;
        $year_filter_sql = '';

        if ($year_filter !== 'all' && $year_filter >= 2000) {
            $year_filter_sql = $wpdb->prepare(' AND YEAR(doc_date) = %d ', $year_filter);
        }

        $queries = [];

        if ($type_filter === 'all' || $type_filter === 'sale') {
            $queries[] = "
                SELECT
                    e.entry_date as doc_date,
                    COALESCE(NULLIF(e.invoice_number_raw,''), '-') as doc_number,
                    COALESCE(NULLIF(e.description,''), 'Factura de venta importada') as party_name,
                    'sale' as doc_type,
                    'imported' as source_type,
                    COALESCE(NULLIF(e.business_line,''), 'other') as business_line,
                    CAST(COALESCE(e.base_amount,0) AS DECIMAL(18,2)) as base_amount,
                    CAST(COALESCE(e.total_amount,0) AS DECIMAL(18,2)) as total_amount,
                    '' as status,
                    '' as paid_at
                FROM {$entries_table} e
                LEFT JOIN {$invoices_table} i2 ON UPPER(TRIM(i2.code)) = UPPER(TRIM(e.invoice_number_normalized))
                WHERE e.entry_group = 'income' AND i2.id IS NULL
            ";

            $queries[] = "
                SELECT
                    i.issue_date as doc_date,
                    COALESCE(NULLIF(i.code,''), CONCAT(i.series, '-', i.number)) as doc_number,
                    COALESCE(NULLIF(c.name,''), 'Cliente sin nombre') as party_name,
                    'sale' as doc_type,
                    'generated' as source_type,
                    COALESCE(NULLIF(i.invoice_type,''), 'standard') as business_line,
                    CAST(COALESCE(i.taxable_base,0) AS DECIMAL(18,2)) as base_amount,
                    CAST(COALESCE(i.total_amount,0) AS DECIMAL(18,2)) as total_amount,
                    COALESCE(i.status,'') as status,
                    COALESCE(i.paid_at,'') as paid_at
                FROM {$invoices_table} i
                LEFT JOIN {$clients_table} c ON c.id = i.client_id
                WHERE i.invoice_type != 'quote'
            ";
        }

        if ($type_filter === 'all' || $type_filter === 'purchase') {
            $queries[] = "
                SELECT
                    e.entry_date as doc_date,
                    COALESCE(NULLIF(e.invoice_number_raw,''), '-') as doc_number,
                    COALESCE(NULLIF(e.description,''), 'Factura de compra importada') as party_name,
                    'purchase' as doc_type,
                    'imported' as source_type,
                    COALESCE(NULLIF(e.account_code,''), 'expense') as business_line,
                    CAST(COALESCE(e.base_amount,0) AS DECIMAL(18,2)) as base_amount,
                    CAST(COALESCE(e.total_amount,0) AS DECIMAL(18,2)) as total_amount,
                    '' as status,
                    '' as paid_at
                FROM {$entries_table} e
                WHERE e.entry_group = 'expense'
            ";
        }

        if (empty($queries)) {
            return [];
        }

        $union_sql = implode(" UNION ALL ", $queries);
        $sql = "SELECT * FROM ({$union_sql}) docs WHERE 1=1 {$year_filter_sql} ORDER BY doc_date DESC, doc_number DESC LIMIT 1000";
        return $wpdb->get_results($sql, ARRAY_A);
    }


    public static function page_analysis() {
        global $wpdb;
        $entries_table = self::accounting_entries_table();
        $imports_table = self::imports_table();
        $selected_month = sanitize_text_field($_GET['forecast_month'] ?? current_time('Y-m'));
        if (!preg_match('/^\d{4}-\d{2}$/', $selected_month)) {
            $selected_month = current_time('Y-m');
        }
        $selected_ts = strtotime($selected_month . '-01');
        if (!$selected_ts) {
            $selected_month = current_time('Y-m');
            $selected_ts = strtotime($selected_month . '-01');
        }
        $prev_month = gmdate('Y-m', strtotime('-1 month', $selected_ts));
        $next_nav_month = gmdate('Y-m', strtotime('+1 month', $selected_ts));
        $selected_month_label = date_i18n('F Y', $selected_ts);
        $prev_url = add_query_arg(['page' => 'vvd_analisis', 'forecast_month' => $prev_month], admin_url('admin.php'));
        $next_url = add_query_arg(['page' => 'vvd_analisis', 'forecast_month' => $next_nav_month], admin_url('admin.php'));
        $today_month = current_time('Y-m');
        $today_url = add_query_arg(['page' => 'vvd_analisis', 'forecast_month' => $today_month], admin_url('admin.php'));

        $available_years = self::analysis_available_years();
        $requested_year = (int) ($_GET['selected_year'] ?? 0);
        $selected_year = $requested_year >= 2000 ? $requested_year : (int) gmdate('Y', $selected_ts);
        if (!in_array($selected_year, $available_years, true) && !empty($available_years)) {
            $selected_year = (int) $available_years[0];
        }
        $monthly_series = self::analysis_monthly_series($selected_year);
        $selected_year_totals = self::totals_from_monthly_series($monthly_series);
        $yearly_series = self::analysis_yearly_history(max(8, count($available_years) + 1));
        $current_year_index = array_search($selected_year, $available_years, true);
        $prev_year = ($current_year_index !== false && isset($available_years[$current_year_index + 1])) ? (int) $available_years[$current_year_index + 1] : $selected_year;
        $next_year = ($current_year_index !== false && $current_year_index > 0 && isset($available_years[$current_year_index - 1])) ? (int) $available_years[$current_year_index - 1] : $selected_year;
        $doc_filter = sanitize_text_field($_GET['doc_filter'] ?? 'all');
        if (!in_array($doc_filter, ['all', 'sale', 'purchase'], true)) {
            $doc_filter = 'all';
        }
        $doc_year = sanitize_text_field($_GET['doc_year'] ?? (string) $selected_year);
        if ($doc_year !== 'all') {
            $doc_year = preg_match('/^\d{4}$/', $doc_year) ? $doc_year : (string) $selected_year;
        }
        $document_rows = self::analysis_documents_list($doc_filter, $doc_year);
        $generated_income_total = (float) $wpdb->get_var("SELECT COALESCE(SUM(taxable_base),0) FROM " . self::table('vvd_invoices') . " WHERE invoice_type != 'quote'");
        $imported_income_total = (float) $wpdb->get_var("SELECT COALESCE(SUM(base_amount),0) FROM {$entries_table} WHERE entry_group='income'");
        $kpi = [
            'income' => $imported_income_total + $generated_income_total,
            'expense' => (float) $wpdb->get_var("SELECT COALESCE(SUM(base_amount),0) FROM {$entries_table} WHERE entry_group='expense'"),
            'income_selected_year' => (float) $selected_year_totals['income'],
            'expense_selected_year' => (float) $selected_year_totals['expense'],
            'agency' => (float) $wpdb->get_var("SELECT COALESCE(SUM(base_amount),0) FROM {$entries_table} WHERE business_line='agency' AND entry_group='income'"),
            'kit' => (float) $wpdb->get_var("SELECT COALESCE(SUM(base_amount),0) FROM {$entries_table} WHERE business_line='kit_digital' AND entry_group='income'"),
            'rectificative' => (float) $wpdb->get_var("SELECT COALESCE(SUM(base_amount),0) FROM {$entries_table} WHERE document_type='rectificative' AND entry_group='income'"),
            'imports' => (int) $wpdb->get_var("SELECT COUNT(*) FROM {$imports_table}"),
        ];
        $net = $kpi['income'] - $kpi['expense'];
        $recurring = self::detect_recurring_expenses(36, $selected_month);
        $recurring_monthly = is_array($recurring) && isset($recurring['monthly']) && is_array($recurring['monthly']) ? $recurring['monthly'] : self::build_empty_monthly_summary($selected_month, gmdate('Y-m', strtotime('+1 month', strtotime($selected_month . '-01'))));
        $recurring_annual = is_array($recurring) && isset($recurring['annual']) && is_array($recurring['annual']) ? $recurring['annual'] : self::build_empty_annual_summary((int) substr($selected_month, 0, 4), (int) substr($selected_month, 5, 2));
        $current_effective_balance = (float) $kpi['income'] - (float) $kpi['expense'];
        $next_month_forecast = (float) $recurring_monthly['next_expected_total'] + (float) $recurring_annual['next_expected_total'];
        $selected_month_total_forecast = (float) $recurring_monthly['current_expected_total'] + (float) $recurring_annual['current_expected_total'];

        echo '<div class="wrap vvd-wrap">';
        echo '<div class="vvd-hero"><h1>Área de análisis</h1><p>Aquí sí: gráficos dentro del plugin, histórico de facturación y previsión del mes que marques. Menos tabla soviética, más lectura visual.</p></div>';

        $prev_year_url = add_query_arg(['page' => 'vvd_analisis', 'forecast_month' => $selected_month, 'selected_year' => $prev_year, 'doc_filter' => $doc_filter, 'doc_year' => $doc_year], admin_url('admin.php'));
        $next_year_url = add_query_arg(['page' => 'vvd_analisis', 'forecast_month' => $selected_month, 'selected_year' => $next_year, 'doc_filter' => $doc_filter, 'doc_year' => $doc_year], admin_url('admin.php'));
        echo '<div class="vvd-chart-grid">';
        echo '<div><div style="display:flex;justify-content:flex-end;gap:8px;align-items:center;margin:0 0 10px"><a class="button" href="' . esc_url($prev_year_url) . '">← Año anterior</a><span class="vvd-pill">Ejercicio ' . esc_html((string) $selected_year) . '</span><a class="button" href="' . esc_url($next_year_url) . '">Año siguiente →</a></div>' . self::render_analysis_month_chart($monthly_series, $selected_year) . '</div>';
        echo '<div class="vvd-balance-card">';
        echo '<div class="vvd-chart-header"><div><h2>Balance y previsión</h2><p>Resumen rápido del cuadro de mando</p></div><span class="vvd-pill">' . esc_html($selected_month_label) . '</span></div>';
        echo '<div class="vvd-balance-big">' . esc_html(($current_effective_balance >= 0 ? '+' : '') . number_format($current_effective_balance, 2, ',', '.')) . ' €</div>';
        echo '<div class="vvd-balance-list">';
        echo '<div class="vvd-balance-item"><span>Ingresos netos acumulados</span><strong class="income">' . esc_html(number_format((float) $kpi['income'], 2, ',', '.')) . ' €</strong></div>';
        echo '<div class="vvd-balance-item"><span>Gastos netos acumulados</span><strong class="expense">' . esc_html(number_format((float) $kpi['expense'], 2, ',', '.')) . ' €</strong></div>';
        echo '<div class="vvd-balance-item"><span>Ingresos ' . esc_html((string) $selected_year) . '</span><strong class="income">' . esc_html(number_format((float) $kpi['income_selected_year'], 2, ',', '.')) . ' €</strong></div>';
        echo '<div class="vvd-balance-item"><span>Gastos ' . esc_html((string) $selected_year) . '</span><strong class="expense">' . esc_html(number_format((float) $kpi['expense_selected_year'], 2, ',', '.')) . ' €</strong></div>';
echo '<div class="vvd-balance-item"><span>Previsión ' . esc_html($selected_month_label) . '</span><strong>' . esc_html(number_format($selected_month_total_forecast, 2, ',', '.')) . ' €</strong></div>';
        echo '<div class="vvd-balance-item"><span>Previsión mes siguiente</span><strong>' . esc_html(number_format($next_month_forecast, 2, ',', '.')) . ' €</strong></div>';
        echo '<div class="vvd-balance-item"><span>Recurrentes detectadas</span><strong>' . esc_html((string) ($recurring_monthly['count'] + $recurring_annual['count'])) . '</strong></div>';
        echo '</div>';
        echo '<p><a class="vvd-btn" href="' . esc_url(admin_url('admin.php?page=vvd_importar_csv')) . '">+ Importar CSV contable</a></p>';
        echo '</div>';
        echo '</div>';

        echo self::render_analysis_history_chart($yearly_series, (float) $kpi['income_selected_year'], $selected_year);

        echo '<div class="vvd-card"><div style="display:flex;justify-content:space-between;align-items:center;gap:12px;flex-wrap:wrap"><div><h2 style="margin:0">Previsión de caja en ' . esc_html($selected_month_label) . '</h2><p class="vvd-muted" style="margin:6px 0 0">Mensuales y anuales integrados en el mismo bloque para ver de un vistazo qué te cae este mes.</p></div><div style="display:flex;gap:8px;align-items:center;flex-wrap:wrap"><a class="button" href="' . esc_url($prev_url) . '">← Mes anterior</a><span class="vvd-pill">' . esc_html($selected_month_label) . '</span><a class="button" href="' . esc_url($next_url) . '">Mes siguiente →</a><a class="button button-secondary" href="' . esc_url($today_url) . '">Volver al mes actual</a></div></div><div class="vvd-kpis">';
        echo '<div class="vvd-kpi"><span>Total estimado del mes</span><strong>' . esc_html(number_format((float) $selected_month_total_forecast, 2, ',', '.')) . ' €</strong></div>';
        echo '<div class="vvd-kpi"><span>Total pendiente del mes</span><strong>' . esc_html(number_format((float) $recurring_monthly['current_pending_total'] + (float) $recurring_annual['current_pending_total'], 2, ',', '.')) . ' €</strong></div>';
        echo '<div class="vvd-kpi"><span>Mensuales detectados</span><strong>' . esc_html((string) count($recurring_monthly['items'])) . '</strong></div>';
        echo '<div class="vvd-kpi"><span>Anuales detectados</span><strong>' . esc_html((string) count($recurring_annual['items'])) . '</strong></div>';
        echo '</div>';
        echo '<table class="vvd-table"><thead><tr><th>Tipo</th><th>Proveedor / concepto</th><th>Patrón</th><th>Día aprox.</th><th>Importe estimado</th><th>Registrado</th><th>Pendiente</th><th>Siguiente referencia</th></tr></thead><tbody>';
        foreach (array_slice($recurring_monthly['items'], 0, 20) as $item) {
            echo '<tr><td><span class="vvd-pill">Mensual</span></td><td><strong>' . esc_html($item['label']) . '</strong><br><span class="vvd-pill">' . esc_html($item['nif'] ?: $item['account_code']) . '</span></td><td>' . esc_html((string) $item['months_count']) . ' meses · último ' . esc_html($item['last_seen_month']) . '</td><td>Día ' . esc_html((string) $item['expected_day']) . '</td><td><strong>' . esc_html(number_format((float) $item['estimated_amount'], 2, ',', '.')) . ' €</strong></td><td>' . esc_html(number_format((float) $item['current_recorded'], 2, ',', '.')) . ' €</td><td>' . esc_html(number_format((float) $item['current_pending'], 2, ',', '.')) . ' €</td><td>' . esc_html($next_nav_month) . '</td></tr>';
        }
        foreach (array_slice($recurring_annual['items'], 0, 20) as $item) {
            echo '<tr><td><span class="vvd-pill">Anual</span></td><td><strong>' . esc_html($item['label']) . '</strong><br><span class="vvd-pill">' . esc_html($item['nif'] ?: $item['account_code']) . '</span></td><td>' . esc_html((string) $item['years_count']) . ' años · último ' . esc_html($item['last_seen_month']) . '</td><td>Día ' . esc_html((string) $item['expected_day']) . '</td><td><strong>' . esc_html(number_format((float) $item['estimated_amount'], 2, ',', '.')) . ' €</strong></td><td>—</td><td>' . esc_html(number_format((float) $item['current_pending'], 2, ',', '.')) . ' €</td><td>' . esc_html(number_format((float) $item['next_expected'], 2, ',', '.')) . ' € próximo año</td></tr>';
        }
        if (empty($recurring_monthly['items']) && empty($recurring_annual['items'])) {
            echo '<tr><td colspan="8">No hay pagos recurrentes detectados para este mes con el histórico cargado.</td></tr>';
        }
        echo '</tbody></table>';
        echo '<p class="vvd-note">Las facturas históricas se guardan en una capa analítica separada para no mezclar la operativa actual con lo heredado.</p></div></div>';
    }

    public static function page_invoice_book() {
        $available_years = self::analysis_available_years();
        $doc_filter = sanitize_text_field($_GET['doc_filter'] ?? 'all');
        if (!in_array($doc_filter, ['all', 'sale', 'purchase'], true)) {
            $doc_filter = 'all';
        }
        $default_year = !empty($available_years) ? (string) $available_years[0] : current_time('Y');
        $doc_year = sanitize_text_field($_GET['doc_year'] ?? $default_year);
        if ($doc_year !== 'all') {
            $doc_year = preg_match('/^\d{4}$/', $doc_year) ? $doc_year : $default_year;
        }
        $doc_client = sanitize_text_field($_GET['doc_client'] ?? '');
        $doc_paid = sanitize_text_field($_GET['doc_paid'] ?? 'all');
        if (!in_array($doc_paid, ['all', 'paid', 'unpaid'], true)) {
            $doc_paid = 'all';
        }
        $document_rows = self::analysis_documents_list($doc_filter, $doc_year);
        if ($doc_client !== '') {
            $needle = function_exists('mb_strtolower') ? mb_strtolower($doc_client) : strtolower($doc_client);
            $document_rows = array_values(array_filter((array) $document_rows, function($doc) use ($needle) {
                $hay = (string) ($doc['party_name'] ?? '');
                $hay = function_exists('mb_strtolower') ? mb_strtolower($hay) : strtolower($hay);
                return strpos($hay, $needle) !== false;
            }));
        }
        if ($doc_paid !== 'all') {
            $document_rows = array_values(array_filter((array) $document_rows, function($doc) use ($doc_paid) {
                if (($doc['doc_type'] ?? '') === 'purchase') {
                    return $doc_paid === 'all';
                }
                $is_paid = !empty($doc['paid_at']) || (($doc['status'] ?? '') === 'paid');
                return $doc_paid === 'paid' ? $is_paid : !$is_paid;
            }));
        }

        echo '<div class="wrap vvd-wrap">';
        echo '<div class="vvd-hero"><h1>Libro de facturas</h1><p>Listado filtrable de compras y ventas. Aquí vive el detalle; en Análisis, los gráficos.</p></div>';
        echo '<div class="vvd-card">';
        echo '<div style="display:flex;justify-content:space-between;align-items:center;gap:12px;flex-wrap:wrap"><div><h2 style="margin:0">Facturas de compra y de venta</h2><p class="vvd-muted" style="margin:6px 0 0">Ventas = históricas importadas + facturas generadas. Compras = gastos importados.</p></div><div class="vvd-actions"><a class="vvd-btn light" href="' . esc_url(admin_url('admin.php?page=vvd_analisis')) . '">Ver análisis</a></div></div>';
        echo '<form method="get" class="vvd-filter-row">';
        echo '<input type="hidden" name="page" value="vvd_libro_facturas">';
        echo '<div class="vvd-field"><label class="vvd-label">Ver</label><select name="doc_filter" class="vvd-select"><option value="all"' . selected($doc_filter, 'all', false) . '>Todas</option><option value="sale"' . selected($doc_filter, 'sale', false) . '>Solo ventas</option><option value="purchase"' . selected($doc_filter, 'purchase', false) . '>Solo compras</option></select></div>';
        echo '<div class="vvd-field"><label class="vvd-label">Año</label><select name="doc_year" class="vvd-select"><option value="all"' . selected($doc_year, 'all', false) . '>Todos los años</option>';
        foreach ($available_years as $yy_option) { echo '<option value="' . esc_attr((string) $yy_option) . '"' . selected($doc_year, (string) $yy_option, false) . '>' . esc_html((string) $yy_option) . '</option>'; }
        echo '</select></div>';
        echo '<div class="vvd-field"><label class="vvd-label">Cliente / proveedor</label><input type="text" name="doc_client" class="vvd-input" value="' . esc_attr($doc_client) . '" placeholder="Nombre del cliente o proveedor"></div>';
        echo '<div class="vvd-field"><label class="vvd-label">Pagadas</label><select name="doc_paid" class="vvd-select"><option value="all"' . selected($doc_paid, 'all', false) . '>Todas</option><option value="paid"' . selected($doc_paid, 'paid', false) . '>Solo pagadas</option><option value="unpaid"' . selected($doc_paid, 'unpaid', false) . '>Solo no pagadas</option></select></div>';
        echo '<div class="vvd-field" style="flex:0 0 auto"><label class="vvd-label">&nbsp;</label><button class="vvd-btn" type="submit">Filtrar</button></div>';
        echo '</form>';
        echo '<table class="vvd-table"><thead><tr><th>Fecha</th><th>Número</th><th>Cliente / proveedor</th><th>Tipo</th><th>Origen</th><th>Pagada</th><th>Línea / serie</th><th>Base</th><th>Total</th></tr></thead><tbody>';
        foreach ((array) $document_rows as $doc) {
            $type_label = ($doc['doc_type'] === 'purchase') ? 'Compra' : 'Venta';
            $type_class = ($doc['doc_type'] === 'purchase') ? 'vvd-type-purchase' : 'vvd-type-sale';
            $source_label = ($doc['source_type'] === 'generated') ? 'Generada' : 'Importada';
            $is_paid = !empty($doc['paid_at']) || (($doc['status'] ?? '') === 'paid');
            echo '<tr>';
            echo '<td>' . esc_html((string) $doc['doc_date']) . '</td>';
            echo '<td><strong>' . esc_html((string) $doc['doc_number']) . '</strong></td>';
            echo '<td>' . esc_html((string) $doc['party_name']) . '</td>';
            echo '<td><span class="' . esc_attr($type_class) . '">' . esc_html($type_label) . '</span></td>';
            echo '<td>' . esc_html($source_label) . '</td>';
            echo '<td>' . (($doc['doc_type'] === 'purchase') ? '<span class="vvd-pill" style="background:#f3f4f6;color:#4b5563">n/d</span>' : ($is_paid ? '<span class="vvd-pill" style="background:#dcfce7;color:#166534">Sí</span>' : '<span class="vvd-pill" style="background:#fef2f2;color:#991b1b">No</span>')) . '</td>';
            echo '<td><span class="vvd-pill">' . esc_html((string) $doc['business_line']) . '</span></td>';
            echo '<td>' . esc_html(number_format((float) $doc['base_amount'], 2, ',', '.')) . ' €</td>';
            echo '<td><strong>' . esc_html(number_format((float) $doc['total_amount'], 2, ',', '.')) . ' €</strong></td>';
            echo '</tr>';
        }
        if (empty($document_rows)) {
            echo '<tr><td colspan="9">No hay facturas para el filtro seleccionado.</td></tr>';
        }
        echo '</tbody></table>';
        echo '<p class="vvd-muted" style="margin-top:10px">Mostrando hasta 1.000 documentos ordenados por fecha descendente.</p>';
        echo '</div></div>';
    }


    protected static function normalize_legacy_status($value) {
        $value = strtoupper(trim((string) $value));
        if ($value === '') {
            return 'issued';
        }
        if (in_array($value, ['CLOSED', 'PAID'], true)) {
            return 'paid';
        }
        if (in_array($value, ['SENT'], true)) {
            return 'sent';
        }
        if (in_array($value, ['DRAFT', 'BORRADOR'], true)) {
            return 'draft';
        }
        return 'issued';
    }

    protected static function extract_xlsx_rows($path) {
        if (!class_exists('ZipArchive')) {
            return ['headers' => [], 'rows' => [], 'error' => 'ZipArchive no está disponible en este servidor.'];
        }
        $zip = new ZipArchive();
        if ($zip->open($path) !== true) {
            return ['headers' => [], 'rows' => [], 'error' => 'No se ha podido abrir el XLSX.'];
        }
        $shared = [];
        $sharedXml = $zip->getFromName('xl/sharedStrings.xml');
        if ($sharedXml) {
            $sx = @simplexml_load_string($sharedXml);
            if ($sx && isset($sx->si)) {
                foreach ($sx->si as $si) {
                    if (isset($si->t)) {
                        $shared[] = (string) $si->t;
                    } else {
                        $parts = [];
                        foreach ($si->r as $run) {
                            $parts[] = (string) $run->t;
                        }
                        $shared[] = implode('', $parts);
                    }
                }
            }
        }
        $sheetXml = $zip->getFromName('xl/worksheets/sheet1.xml');
        if (!$sheetXml) {
            $zip->close();
            return ['headers' => [], 'rows' => [], 'error' => 'No se ha encontrado la primera hoja del XLSX.'];
        }
        $sheet = @simplexml_load_string($sheetXml);
        if (!$sheet || !isset($sheet->sheetData->row)) {
            $zip->close();
            return ['headers' => [], 'rows' => [], 'error' => 'La hoja del XLSX no tiene datos legibles.'];
        }
        $all_rows = [];
        foreach ($sheet->sheetData->row as $row) {
            $cells = [];
            foreach ($row->c as $cell) {
                $ref = (string) $cell['r'];
                $letters = preg_replace('/\d+/', '', $ref);
                $idx = 0;
                for ($i = 0; $i < strlen($letters); $i++) {
                    $idx = $idx * 26 + (ord($letters[$i]) - 64);
                }
                $value = '';
                $type = (string) $cell['t'];
                if ($type === 'inlineStr' && isset($cell->is->t)) {
                    $value = (string) $cell->is->t;
                } elseif ($type === 's') {
                    $value = $shared[(int) $cell->v] ?? '';
                } else {
                    $value = isset($cell->v) ? (string) $cell->v : '';
                }
                $cells[$idx] = $value;
            }
            if ($cells) {
                ksort($cells);
                $max = max(array_keys($cells));
                $rowValues = [];
                for ($i = 1; $i <= $max; $i++) {
                    $rowValues[] = isset($cells[$i]) ? trim((string) $cells[$i]) : '';
                }
                $all_rows[] = $rowValues;
            }
        }
        $zip->close();
        if (!$all_rows) {
            return ['headers' => [], 'rows' => [], 'error' => 'El XLSX está vacío.'];
        }
        $headers = array_shift($all_rows);
        $headers = array_map(function($h){ return trim((string) $h); }, $headers);
        $assocRows = [];
        foreach ($all_rows as $row) {
            $assoc = [];
            foreach ($headers as $i => $header) {
                $assoc[$header] = isset($row[$i]) ? trim((string) $row[$i]) : '';
            }
            $assocRows[] = $assoc;
        }
        return ['headers' => $headers, 'rows' => $assocRows, 'error' => ''];
    }

    protected static function excel_serial_to_date($value) {
        $value = trim((string) $value);
        if ($value === '') {
            return '';
        }
        if (preg_match('/^\d{4}-\d{2}-\d{2}$/', $value)) {
            return $value;
        }
        if (preg_match('/^\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4}$/', $value)) {
            return self::normalize_csv_date($value);
        }
        if (is_numeric($value)) {
            $ts = ((float) $value - 25569) * 86400;
            if ($ts > 0) {
                return gmdate('Y-m-d', (int) round($ts));
            }
        }
        return '';
    }

    protected static function legacy_invoice_type_from_row($row) {
        $rawNumber = self::normalize_invoice_number($row['Invoice Number'] ?? '');
        $invoiceType = strtoupper(trim((string) ($row['Invoice Type'] ?? '')));
        if (strpos($rawNumber, 'RT-') === 0 || strpos($invoiceType, 'RECT') !== false) {
            return 'rectificativa';
        }
        if (strpos($rawNumber, 'P-') === 0 || strpos($invoiceType, 'QUOTE') !== false || strpos($invoiceType, 'ESTIMATE') !== false) {
            return 'quote';
        }
        return 'standard';
    }

    protected static function legacy_series_and_number($invoice_number, $invoice_type = 'standard') {
        $invoice_number = self::normalize_invoice_number($invoice_number);
        $settings = self::settings();
        $series = $invoice_type === 'rectificativa' ? strtoupper($settings['rectificative_series'] ?: 'RT') : ($invoice_type === 'quote' ? strtoupper($settings['quote_series'] ?: 'P') : strtoupper($settings['default_series'] ?: 'A'));
        if ($invoice_number === '') {
            return ['series' => $series, 'number' => 0, 'code' => null];
        }
        if (preg_match('/^([A-Z]+)-?(\d{1,})$/', $invoice_number, $m)) {
            return ['series' => strtoupper($m[1]), 'number' => (int) $m[2], 'code' => strtoupper($m[1]) . '-' . str_pad((int) $m[2], max(3, strlen($m[2])), '0', STR_PAD_LEFT)];
        }
        if (preg_match('/^(\d{3,})$/', $invoice_number, $m)) {
            return ['series' => 'KD', 'number' => (int) $m[1], 'code' => $invoice_number];
        }
        return ['series' => $series, 'number' => 0, 'code' => $invoice_number];
    }

    protected static function upsert_client_from_legacy_row($row) {
        global $wpdb;
        $clients_table = self::table('vvd_clients');
        $tax_id = sanitize_text_field($row['CF.NIF/CIF'] ?? '');
        $name = sanitize_text_field($row['Customer Name'] ?? '');
        $external_id = sanitize_text_field($row['Customer ID'] ?? '');
        $email = sanitize_email($row['CF.E-mail'] ?? '');
        $phone = sanitize_text_field($row['CF.Phone'] ?? '');
        $mobile = sanitize_text_field($row['CF.Mobile Phone'] ?? '');
        $client_id = 0;
        if ($external_id !== '') {
            $client_id = (int) $wpdb->get_var($wpdb->prepare("SELECT id FROM {$clients_table} WHERE external_id=%s LIMIT 1", $external_id));
        }
        if (!$client_id && $tax_id !== '') {
            $client_id = (int) $wpdb->get_var($wpdb->prepare("SELECT id FROM {$clients_table} WHERE tax_id=%s LIMIT 1", $tax_id));
        }
        if (!$client_id && $name !== '') {
            $client_id = (int) $wpdb->get_var($wpdb->prepare("SELECT id FROM {$clients_table} WHERE name=%s LIMIT 1", $name));
        }
        $payload = [
            'name' => $name !== '' ? $name : 'Cliente histórico',
            'tax_id' => $tax_id !== '' ? $tax_id : 'PENDIENTE',
            'external_id' => $external_id ?: null,
            'address' => sanitize_text_field($row['CF.Billing Street'] ?? '') ?: null,
            'city' => sanitize_text_field($row['CF.Billing City'] ?? '') ?: null,
            'state' => sanitize_text_field($row['CF.Billing State'] ?? '') ?: null,
            'postcode' => sanitize_text_field($row['CF.Billing Zip'] ?? '') ?: null,
            'country' => sanitize_text_field($row['CF.Billing Country'] ?? '') ?: null,
            'email' => $email ?: null,
            'phone' => $phone ?: null,
            'mobile' => $mobile ?: null,
            'updated_at' => current_time('mysql'),
        ];
        if ($client_id) {
            $wpdb->update($clients_table, $payload, ['id' => $client_id]);
            return $client_id;
        }
        $payload['created_at'] = current_time('mysql');
        $wpdb->insert($clients_table, $payload);
        return (int) $wpdb->insert_id;
    }

    protected static function import_legacy_rows_to_invoices($rows, $mode = 'merge') {
        global $wpdb;
        $invoices_table = self::table('vvd_invoices');
        $lines_table = self::table('vvd_invoice_lines');
        $grouped = [];
        foreach ((array) $rows as $row) {
            $number = self::normalize_invoice_number($row['Invoice Number'] ?? '');
            $external_id = trim((string) ($row['Invoice ID'] ?? ''));
            $group_key = $external_id !== '' ? 'id:' . $external_id : 'num:' . $number;
            if ($group_key === 'num:') {
                continue;
            }
            if (!isset($grouped[$group_key])) {
                $grouped[$group_key] = ['header' => $row, 'lines' => []];
            }
            $grouped[$group_key]['lines'][] = $row;
        }
        $stats = ['created' => 0, 'updated' => 0, 'skipped' => 0, 'lines' => 0];
        foreach ($grouped as $bundle) {
            $row = $bundle['header'];
            $invoice_type = self::legacy_invoice_type_from_row($row);
            if ($invoice_type === 'quote') {
                continue;
            }
            $series_info = self::legacy_series_and_number($row['Invoice Number'] ?? '', $invoice_type);
            $client_id = self::upsert_client_from_legacy_row($row);
            $code = $series_info['code'] ?: self::normalize_invoice_number($row['Invoice Number'] ?? '');
            $external_id = sanitize_text_field($row['Invoice ID'] ?? '');
            $existing_id = 0;
            if ($external_id !== '') {
                $existing_id = (int) $wpdb->get_var($wpdb->prepare("SELECT id FROM {$invoices_table} WHERE external_invoice_id=%s LIMIT 1", $external_id));
            }
            if (!$existing_id && $code !== '') {
                $existing_id = (int) $wpdb->get_var($wpdb->prepare("SELECT id FROM {$invoices_table} WHERE code=%s LIMIT 1", $code));
            }
            if ($existing_id && $mode === 'import_only_missing') {
                $stats['skipped']++;
                continue;
            }
            $issue_date = self::excel_serial_to_date($row['Invoice Date'] ?? '') ?: self::excel_serial_to_date($row['Issued Date'] ?? '') ?: current_time('Y-m-d');
            $due_date = self::excel_serial_to_date($row['Due Date'] ?? '');
            $expected_payment_date = self::excel_serial_to_date($row['Expected Payment Date'] ?? '');
            $last_payment_date = self::excel_serial_to_date($row['Last Payment Date'] ?? '');
            $subtotal = self::normalize_amount($row['SubTotal'] ?? 0);
            $total = self::normalize_amount($row['Total'] ?? 0);
            $withholding = self::normalize_amount($row['TotalRetentionAmountFCY'] ?? ($row['TotalRetentionAmountBCY'] ?? 0));
            $tax_amount = max(0, round($total - $subtotal + $withholding, 2));
            $payload = [
                'series' => $series_info['series'],
                'series_label' => 'Histórica ' . $series_info['series'],
                'number' => (int) $series_info['number'],
                'code' => $code ?: null,
                'invoice_type' => $invoice_type,
                'source_origin' => 'legacy_import',
                'external_invoice_id' => $external_id ?: null,
                'external_invoice_number' => sanitize_text_field($row['Invoice Number'] ?? ''),
                'issue_date' => $issue_date,
                'due_date' => $due_date ?: null,
                'client_id' => $client_id,
                'client_external_id' => sanitize_text_field($row['Customer ID'] ?? ''),
                'billing_name' => sanitize_text_field($row['Customer Name'] ?? ''),
                'billing_address' => sanitize_text_field($row['CF.Billing Street'] ?? ''),
                'billing_city' => sanitize_text_field($row['CF.Billing City'] ?? ''),
                'billing_state' => sanitize_text_field($row['CF.Billing State'] ?? ''),
                'billing_postcode' => sanitize_text_field($row['CF.Billing Zip'] ?? ''),
                'billing_country' => sanitize_text_field($row['CF.Billing Country'] ?? ''),
                'client_email' => sanitize_email($row['CF.E-mail'] ?? ''),
                'client_phone' => sanitize_text_field($row['CF.Phone'] ?? ''),
                'client_mobile' => sanitize_text_field($row['CF.Mobile Phone'] ?? ''),
                'notes' => sanitize_textarea_field($row['Notes'] ?? ''),
                'terms_conditions' => sanitize_textarea_field($row['Terms & Conditions'] ?? ''),
                'estimate_number' => sanitize_text_field($row['Estimate Number'] ?? ''),
                'currency_code' => sanitize_text_field($row['Currency Code'] ?? 'EUR'),
                'payment_terms' => sanitize_text_field($row['Payment Terms'] ?? ''),
                'payment_terms_label' => sanitize_text_field($row['Payment Terms Label'] ?? ''),
                'expected_payment_date' => $expected_payment_date ?: null,
                'last_payment_date' => $last_payment_date ?: null,
                'paid_at' => (self::normalize_legacy_status($row['Invoice Status'] ?? '') === 'paid') ? ($last_payment_date ?: current_time('mysql')) : null,
                'taxable_base' => $subtotal,
                'subtotal' => $subtotal,
                'discount_total' => self::normalize_amount($row['Entity Discount Amount'] ?? 0),
                'tax_rate' => 0,
                'tax_amount' => $tax_amount,
                'irpf_rate' => 0,
                'irpf_amount' => $withholding,
                'withholding_total' => $withholding,
                'total_amount' => $total,
                'balance_due' => self::normalize_amount($row['Balance'] ?? 0),
                'status' => self::normalize_legacy_status($row['Invoice Status'] ?? ''),
                'legacy_status' => sanitize_text_field($row['Invoice Status'] ?? ''),
                'verifactu_status' => 'not-applicable',
                'updated_at' => current_time('mysql'),
                'import_hash' => md5(($external_id ?: $code ?: wp_json_encode($row)) . '|' . $issue_date),
            ];
            if ($existing_id) {
                $wpdb->update($invoices_table, $payload, ['id' => $existing_id]);
                $invoice_id = $existing_id;
                $stats['updated']++;
            } else {
                $payload['created_at'] = current_time('mysql');
                $wpdb->insert($invoices_table, $payload);
                $invoice_id = (int) $wpdb->insert_id;
                $stats['created']++;
            }
            if ($invoice_id) {
                $wpdb->delete($lines_table, ['invoice_id' => $invoice_id]);
                $sort_order = 0;
                foreach ((array) $bundle['lines'] as $line) {
                    $item_name = sanitize_text_field($line['Item Name'] ?? '');
                    $item_desc = sanitize_textarea_field($line['Item Desc'] ?? '');
                    $quantity = self::normalize_amount($line['Quantity'] ?? 0);
                    $line_total = self::normalize_amount($line['Item Total'] ?? 0);
                    $unit_price = $quantity > 0 ? round($line_total / $quantity, 6) : $line_total;
                    $wpdb->insert($lines_table, [
                        'invoice_id' => $invoice_id,
                        'item_name' => $item_name ?: null,
                        'item_desc' => $item_desc ?: null,
                        'description' => $item_name !== '' ? $item_name : ($item_desc !== '' ? mb_substr($item_desc, 0, 250) : 'Concepto histórico'),
                        'quantity' => $quantity > 0 ? $quantity : 1,
                        'discount_amount' => self::normalize_amount($line['Discount Amount'] ?? 0),
                        'unit_price' => $unit_price,
                        'line_total' => $line_total,
                        'usage_unit' => sanitize_text_field($line['Usage unit'] ?? ''),
                        'sort_order' => $sort_order++,
                    ]);
                    $stats['lines']++;
                }
            }
        }
        return $stats;
    }


    protected static function contact_row_first_non_empty($row, $keys) {
        foreach ((array) $keys as $key) {
            if (!array_key_exists($key, (array) $row)) {
                continue;
            }
            $value = trim((string) $row[$key]);
            if ($value !== '') {
                return $value;
            }
        }
        return '';
    }

    protected static function client_import_display_name($row) {
        $company = self::contact_row_first_non_empty($row, ['Company Name', 'Display Name', 'Contact Name']);
        if ($company !== '') {
            return sanitize_text_field($company);
        }
        $full = trim(self::contact_row_first_non_empty($row, ['First Name']) . ' ' . self::contact_row_first_non_empty($row, ['Last Name']));
        if ($full !== '') {
            return sanitize_text_field($full);
        }
        return 'Cliente importado';
    }

    protected static function import_contact_rows_to_clients($rows, $mode = 'merge') {
        global $wpdb;
        $clients_table = self::table('vvd_clients');
        $stats = ['created' => 0, 'updated' => 0, 'skipped' => 0, 'linked' => 0];

        foreach ((array) $rows as $row) {
            $name = self::client_import_display_name($row);
            $tax_id = sanitize_text_field(self::contact_row_first_non_empty($row, ['CF.NIF/CIF', 'TaxID', 'SIRET']));
            $email = sanitize_email(self::contact_row_first_non_empty($row, ['EmailID']));
            $external_id = sanitize_text_field(self::contact_row_first_non_empty($row, ['Contact ID', 'ID de empresa']));
            $phone = sanitize_text_field(self::contact_row_first_non_empty($row, ['Phone', 'Billing Phone']));
            $mobile = sanitize_text_field(self::contact_row_first_non_empty($row, ['MobilePhone']));
            $address = trim(self::contact_row_first_non_empty($row, ['Billing Address']) . ' ' . self::contact_row_first_non_empty($row, ['Billing Street2']));
            $city = sanitize_text_field(self::contact_row_first_non_empty($row, ['Billing City']));
            $state = sanitize_text_field(self::contact_row_first_non_empty($row, ['Billing State']));
            $postcode = sanitize_text_field(self::contact_row_first_non_empty($row, ['Billing Code']));
            $country = sanitize_text_field(self::contact_row_first_non_empty($row, ['Billing Country']));

            if ($name === '' && $tax_id === '' && $email === '') {
                $stats['skipped']++;
                continue;
            }

            $client_id = 0;
            if ($external_id !== '') {
                $client_id = (int) $wpdb->get_var($wpdb->prepare("SELECT id FROM {$clients_table} WHERE external_id=%s LIMIT 1", $external_id));
            }
            if (!$client_id && $tax_id !== '') {
                $client_id = (int) $wpdb->get_var($wpdb->prepare("SELECT id FROM {$clients_table} WHERE tax_id=%s LIMIT 1", $tax_id));
            }
            if (!$client_id && $email !== '') {
                $client_id = (int) $wpdb->get_var($wpdb->prepare("SELECT id FROM {$clients_table} WHERE email=%s LIMIT 1", $email));
            }
            if (!$client_id && $name !== '') {
                $client_id = (int) $wpdb->get_var($wpdb->prepare("SELECT id FROM {$clients_table} WHERE name=%s LIMIT 1", $name));
            }

            if ($client_id && $mode === 'import_only_missing') {
                $stats['skipped']++;
                continue;
            }

            $contact_payload = [
                'id' => $external_id,
                'email' => $email,
                'name' => sanitize_text_field(self::contact_row_first_non_empty($row, ['Contact Name'])) ?: trim(self::contact_row_first_non_empty($row, ['First Name']) . ' ' . self::contact_row_first_non_empty($row, ['Last Name'])),
                'salutation' => sanitize_text_field(self::contact_row_first_non_empty($row, ['Salutation'])),
                'first_name' => sanitize_text_field(self::contact_row_first_non_empty($row, ['First Name'])),
                'last_name' => sanitize_text_field(self::contact_row_first_non_empty($row, ['Last Name'])),
                'phone' => $phone,
                'mobile' => $mobile,
                'designation' => sanitize_text_field(self::contact_row_first_non_empty($row, ['Designation'])),
                'department' => sanitize_text_field(self::contact_row_first_non_empty($row, ['Department'])),
                'contact_type' => sanitize_text_field(self::contact_row_first_non_empty($row, ['Contact Type'])),
                'is_primary' => self::contact_row_first_non_empty($row, ['Primary Contact ID']) === $external_id,
            ];

            $payload = [
                'name' => $name,
                'display_name' => sanitize_text_field(self::contact_row_first_non_empty($row, ['Display Name'])),
                'company_name' => sanitize_text_field(self::contact_row_first_non_empty($row, ['Company Name'])),
                'salutation' => $contact_payload['salutation'] ?: null,
                'first_name' => $contact_payload['first_name'] ?: null,
                'last_name' => $contact_payload['last_name'] ?: null,
                'contact_name' => $contact_payload['name'] ?: null,
                'contact_type' => $contact_payload['contact_type'] ?: null,
                'designation' => $contact_payload['designation'] ?: null,
                'department' => $contact_payload['department'] ?: null,
                'tax_id' => $tax_id,
                'external_id' => $external_id ?: null,
                'contact_id_legacy' => $external_id ?: null,
                'primary_contact_id' => sanitize_text_field(self::contact_row_first_non_empty($row, ['Primary Contact ID'])) ?: null,
                'address' => sanitize_text_field($address),
                'billing_attention' => sanitize_text_field(self::contact_row_first_non_empty($row, ['Billing Attention'])) ?: null,
                'billing_street2' => sanitize_text_field(self::contact_row_first_non_empty($row, ['Billing Street2'])) ?: null,
                'city' => $city ?: null,
                'state' => $state ?: null,
                'postcode' => $postcode ?: null,
                'country' => $country ?: null,
                'shipping_attention' => sanitize_text_field(self::contact_row_first_non_empty($row, ['Shipping Attention'])) ?: null,
                'shipping_address' => sanitize_text_field(trim(self::contact_row_first_non_empty($row, ['Shipping Address']) . ' ' . self::contact_row_first_non_empty($row, ['Shipping Street2']))) ?: null,
                'shipping_street2' => sanitize_text_field(self::contact_row_first_non_empty($row, ['Shipping Street2'])) ?: null,
                'shipping_city' => sanitize_text_field(self::contact_row_first_non_empty($row, ['Shipping City'])) ?: null,
                'shipping_state' => sanitize_text_field(self::contact_row_first_non_empty($row, ['Shipping State'])) ?: null,
                'shipping_postcode' => sanitize_text_field(self::contact_row_first_non_empty($row, ['Shipping Code'])) ?: null,
                'shipping_country' => sanitize_text_field(self::contact_row_first_non_empty($row, ['Shipping Country'])) ?: null,
                'email' => $email ?: null,
                'phone' => $phone ?: null,
                'mobile' => $mobile ?: null,
                'website' => esc_url_raw(self::contact_row_first_non_empty($row, ['Website'])),
                'client_status' => sanitize_text_field(self::contact_row_first_non_empty($row, ['Status'])) ?: null,
                'currency_code' => sanitize_text_field(self::contact_row_first_non_empty($row, ['Currency Code'])) ?: null,
                'notes' => sanitize_textarea_field(self::contact_row_first_non_empty($row, ['Notes'])) ?: null,
                'payment_terms' => sanitize_text_field(self::contact_row_first_non_empty($row, ['Payment Terms'])) ?: null,
                'payment_terms_label' => sanitize_text_field(self::contact_row_first_non_empty($row, ['Payment Terms Label'])) ?: null,
                'taxable' => sanitize_text_field(self::contact_row_first_non_empty($row, ['Taxable'])) ?: null,
                'tax_name' => sanitize_text_field(self::contact_row_first_non_empty($row, ['Tax Name'])) ?: null,
                'tax_percentage' => sanitize_text_field(self::contact_row_first_non_empty($row, ['Tax Percentage'])) ?: null,
                'exemption_reason' => sanitize_textarea_field(self::contact_row_first_non_empty($row, ['Exemption Reason'])) ?: null,
                'contacts_json' => self::append_client_contact('', $contact_payload),
                'raw_import_json' => wp_json_encode($row),
                'updated_at' => current_time('mysql'),
            ];

            if ($client_id) {
                $existing = (array) $wpdb->get_row($wpdb->prepare("SELECT * FROM {$clients_table} WHERE id=%d", $client_id), ARRAY_A);
                foreach ($payload as $key => $value) {
                    if (in_array($key, ['contacts_json', 'raw_import_json'], true)) {
                        continue;
                    }
                    if ($value === null || $value === '') {
                        $payload[$key] = $existing[$key] ?? $value;
                    }
                }
                $payload['contacts_json'] = self::append_client_contact($existing['contacts_json'] ?? '', $contact_payload);
                $payload['raw_import_json'] = wp_json_encode($row);
                $wpdb->update($clients_table, $payload, ['id' => $client_id]);
                $stats['updated']++;
            } else {
                $payload['created_at'] = current_time('mysql');
                $wpdb->insert($clients_table, $payload);
                $client_id = (int) $wpdb->insert_id;
                $stats['created']++;
            }

            if ($client_id) {
                $linked = $wpdb->query($wpdb->prepare(
                    "UPDATE " . self::table('vvd_invoices') . " 
                     SET client_id=%d,
                         client_external_id=CASE WHEN (client_external_id IS NULL OR client_external_id='') THEN %s ELSE client_external_id END,
                         billing_name=CASE WHEN (billing_name IS NULL OR billing_name='') THEN %s ELSE billing_name END,
                         client_email=CASE WHEN (client_email IS NULL OR client_email='') THEN %s ELSE client_email END,
                         client_phone=CASE WHEN (client_phone IS NULL OR client_phone='') THEN %s ELSE client_phone END,
                         client_mobile=CASE WHEN (client_mobile IS NULL OR client_mobile='') THEN %s ELSE client_mobile END
                     WHERE (
                        (%s <> '' AND client_external_id=%s)
                        OR (%s <> '' AND billing_name=%s)
                        OR (%s <> '' AND client_email=%s)
                     )",
                    $client_id,
                    $external_id,
                    $name,
                    $email,
                    $phone,
                    $mobile,
                    $external_id, $external_id,
                    $name, $name,
                    $email, $email
                ));
                if ($linked) {
                    $stats['linked'] += (int) $linked;
                }
            }
        }

        return $stats;
    }

    public static function page_import_clients() {
        self::page_import_accounting();
    }

    public static function import_clients_xlsx() {
        if (!current_user_can('manage_options')) { wp_die('No autorizado'); }
        check_admin_referer('vvd_import_clients_xlsx');
        if (empty($_FILES['clients_xlsx']['tmp_name'])) {
            wp_safe_redirect(admin_url('admin.php?page=vvd_importar_clientes&vvd_import_status=error&vvd_message=' . rawurlencode('No se ha recibido el XLSX de contactos.')));
            exit;
        }
        $mode = sanitize_text_field($_POST['clients_mode'] ?? 'merge');
        if (!in_array($mode, ['merge', 'import_only_missing'], true)) {
            $mode = 'merge';
        }
        $tmp = $_FILES['clients_xlsx']['tmp_name'];
        $filename = sanitize_file_name($_FILES['clients_xlsx']['name'] ?? 'contactos.xlsx');
        $hash = hash_file('sha256', $tmp);
        global $wpdb;
        $imports_table = self::imports_table();
        $existing = (int) $wpdb->get_var($wpdb->prepare("SELECT id FROM {$imports_table} WHERE hash_file=%s LIMIT 1", $hash));
        if ($existing) {
            wp_safe_redirect(admin_url('admin.php?page=vvd_importar_clientes&vvd_import_status=error&vvd_message=' . rawurlencode('Ese Excel de clientes ya fue importado antes.')));
            exit;
        }
        $parsed = self::extract_xlsx_rows($tmp);
        if (!empty($parsed['error'])) {
            wp_safe_redirect(admin_url('admin.php?page=vvd_importar_clientes&vvd_import_status=error&vvd_message=' . rawurlencode($parsed['error'])));
            exit;
        }
        $stats = self::import_contact_rows_to_clients($parsed['rows'], $mode);
        $wpdb->insert($imports_table, [
            'filename' => $filename,
            'source_type' => 'clients_xlsx_' . $mode,
            'hash_file' => $hash,
            'total_rows' => count((array) $parsed['rows']),
            'valid_rows' => (int) $stats['created'] + (int) $stats['updated'],
            'error_rows' => 0,
            'meta_json' => wp_json_encode($stats),
            'imported_by' => get_current_user_id(),
            'imported_at' => current_time('mysql'),
        ]);
        self::log_event('clients_xlsx_imported', 'Importado Excel de clientes', null, null, $stats);
        $msg = sprintf('Clientes importados. Creados: %d · Actualizados: %d · Omitidos: %d · Facturas enlazadas: %d', (int) $stats['created'], (int) $stats['updated'], (int) $stats['skipped'], (int) $stats['linked']);
        wp_safe_redirect(admin_url('admin.php?page=vvd_importar_clientes&vvd_import_status=ok&vvd_message=' . rawurlencode($msg)));
        exit;
    }

    public static function import_legacy_xlsx() {
        if (!current_user_can('manage_options')) { wp_die('No autorizado'); }
        check_admin_referer('vvd_import_legacy_xlsx');
        if (empty($_FILES['legacy_xlsx']['tmp_name'])) {
            wp_safe_redirect(admin_url('admin.php?page=vvd_importar_csv&vvd_import_status=error&vvd_message=' . rawurlencode('No se ha recibido el XLSX histórico.')));
            exit;
        }
        $mode = sanitize_text_field($_POST['legacy_mode'] ?? 'merge');
        if (!in_array($mode, ['merge', 'import_only_missing'], true)) {
            $mode = 'merge';
        }
        $tmp = $_FILES['legacy_xlsx']['tmp_name'];
        $filename = sanitize_file_name($_FILES['legacy_xlsx']['name'] ?? 'historico.xlsx');
        $hash = hash_file('sha256', $tmp);
        global $wpdb;
        $imports_table = self::imports_table();
        $existing = (int) $wpdb->get_var($wpdb->prepare("SELECT id FROM {$imports_table} WHERE hash_file=%s LIMIT 1", $hash));
        if ($existing) {
            wp_safe_redirect(admin_url('admin.php?page=vvd_importar_csv&vvd_import_status=error&vvd_message=' . rawurlencode('Ese Excel ya fue importado antes.')));
            exit;
        }
        $parsed = self::extract_xlsx_rows($tmp);
        if (!empty($parsed['error'])) {
            wp_safe_redirect(admin_url('admin.php?page=vvd_importar_csv&vvd_import_status=error&vvd_message=' . rawurlencode($parsed['error'])));
            exit;
        }
        $stats = self::import_legacy_rows_to_invoices($parsed['rows'], $mode);
        $wpdb->insert($imports_table, [
            'filename' => $filename,
            'source_type' => 'legacy_xlsx_' . $mode,
            'hash_file' => $hash,
            'total_rows' => count((array) $parsed['rows']),
            'valid_rows' => (int) $stats['created'] + (int) $stats['updated'],
            'error_rows' => 0,
            'meta_json' => wp_json_encode($stats),
            'imported_by' => get_current_user_id(),
            'imported_at' => current_time('mysql'),
        ]);
        self::log_event('legacy_xlsx_imported', 'Importado Excel histórico', null, null, $stats);
        $msg = sprintf('Histórico importado. Creadas: %d · Actualizadas: %d · Omitidas: %d · Líneas: %d', (int) $stats['created'], (int) $stats['updated'], (int) $stats['skipped'], (int) $stats['lines']);
        wp_safe_redirect(admin_url('admin.php?page=vvd_importar_csv&vvd_import_status=ok&vvd_message=' . rawurlencode($msg)));
        exit;
    }


    protected static function imported_clients_where_sql() {
        return "(raw_import_json IS NOT NULL AND raw_import_json <> '') OR (contact_id_legacy IS NOT NULL AND contact_id_legacy <> '') OR (contacts_json IS NOT NULL AND contacts_json <> '')";
    }

    public static function delete_legacy_imports() {
        if (!current_user_can('manage_options')) { wp_die('No autorizado'); }
        check_admin_referer('vvd_delete_legacy_imports');
        global $wpdb;
        $invoices_table = self::table('vvd_invoices');
        $lines_table = self::table('vvd_invoice_lines');
        $events_table = self::table('vvd_events');
        $imports_table = self::imports_table();

        $ids = $wpdb->get_col("SELECT id FROM {$invoices_table} WHERE source_origin='legacy_import'");
        $deleted_lines = 0;
        $deleted_events = 0;
        $deleted_invoices = 0;
        if (!empty($ids)) {
            $ids = array_map('intval', $ids);
            $ids_sql = implode(',', $ids);
            $deleted_lines = (int) $wpdb->query("DELETE FROM {$lines_table} WHERE invoice_id IN ({$ids_sql})");
            $deleted_events = (int) $wpdb->query("DELETE FROM {$events_table} WHERE invoice_id IN ({$ids_sql})");
            $deleted_invoices = (int) $wpdb->query("DELETE FROM {$invoices_table} WHERE id IN ({$ids_sql})");
        }
        $deleted_imports = (int) $wpdb->query($wpdb->prepare("DELETE FROM {$imports_table} WHERE source_type LIKE %s", 'legacy_xlsx_%'));
        self::log_event('legacy_imports_deleted', 'Importaciones históricas eliminadas', null, null, [
            'invoices' => $deleted_invoices,
            'lines' => $deleted_lines,
            'events' => $deleted_events,
            'imports' => $deleted_imports,
        ]);
        $msg = sprintf('Importaciones históricas eliminadas. Facturas: %d · Líneas: %d · Eventos: %d · Registros de importación: %d', $deleted_invoices, $deleted_lines, $deleted_events, $deleted_imports);
        wp_safe_redirect(admin_url('admin.php?page=vvd_importar_csv&vvd_import_status=ok&vvd_message=' . rawurlencode($msg)));
        exit;
    }

    public static function delete_imported_clients() {
        if (!current_user_can('manage_options')) { wp_die('No autorizado'); }
        check_admin_referer('vvd_delete_imported_clients');
        global $wpdb;
        $clients_table = self::table('vvd_clients');
        $events_table = self::table('vvd_events');
        $imports_table = self::imports_table();

        $where = self::imported_clients_where_sql();
        $rows = $wpdb->get_results("SELECT c.id, (SELECT COUNT(*) FROM " . self::table('vvd_invoices') . " i WHERE i.client_id = c.id) AS invoice_count FROM {$clients_table} c WHERE {$where}", ARRAY_A);
        $deletable_ids = [];
        $skipped = 0;
        foreach ((array) $rows as $row) {
            if ((int) $row['invoice_count'] > 0) {
                $skipped++;
            } else {
                $deletable_ids[] = (int) $row['id'];
            }
        }
        $deleted_clients = 0;
        $deleted_events = 0;
        if (!empty($deletable_ids)) {
            $ids_sql = implode(',', $deletable_ids);
            $deleted_events = (int) $wpdb->query("DELETE FROM {$events_table} WHERE client_id IN ({$ids_sql})");
            $deleted_clients = (int) $wpdb->query("DELETE FROM {$clients_table} WHERE id IN ({$ids_sql})");
        }
        $deleted_imports = (int) $wpdb->query($wpdb->prepare("DELETE FROM {$imports_table} WHERE source_type LIKE %s", 'clients_xlsx_%'));
        self::log_event('imported_clients_deleted', 'Clientes importados eliminados', null, null, [
            'clients' => $deleted_clients,
            'events' => $deleted_events,
            'imports' => $deleted_imports,
            'skipped_linked' => $skipped,
        ]);
        $msg = sprintf('Clientes importados eliminados. Clientes borrados: %d · Eventos: %d · Registros de importación: %d · Omitidos por tener documentos asociados: %d', $deleted_clients, $deleted_events, $deleted_imports, $skipped);
        wp_safe_redirect(admin_url('admin.php?page=vvd_importar_clientes&vvd_import_status=ok&vvd_message=' . rawurlencode($msg)));
        exit;
    }

    public static function page_import_accounting() {
        $status = sanitize_text_field($_GET['vvd_import_status'] ?? '');
        $message = sanitize_text_field($_GET['vvd_message'] ?? '');
        echo '<div class="wrap vvd-wrap"><h1>Importaciones</h1>';
        if ($message !== '') {
            $class = $status === 'ok' ? 'updated notice' : 'error notice';
            echo '<div class="' . esc_attr($class) . '"><p>' . esc_html($message) . '</p></div>';
        }
        echo '<div class="vvd-grid">';
        echo '<div class="vvd-col-6"><div class="vvd-card"><h2>Importar CSV contable</h2><p class="vvd-note">Sirve tanto para facturas históricas como para gastos trimestrales. El importador clasifica FAC como agencia, KD o 3 dígitos como Kit Digital y RT como rectificativas.</p><form method="post" action="' . esc_url(admin_url('admin-post.php')) . '" enctype="multipart/form-data">';
        wp_nonce_field('vvd_import_accounting_csv');
        echo '<input type="hidden" name="action" value="vvd_import_accounting_csv">';
        echo '<div class="vvd-grid">';
        echo '<div class="vvd-col-12"><label class="vvd-label">Archivo CSV</label><input class="vvd-input" type="file" name="accounting_csv" accept=".csv,text/csv" required></div>';
        echo '<div class="vvd-col-12"><label class="vvd-label">Tipo orientativo</label><select class="vvd-select" name="source_type"><option value="auto">Autodetectar</option><option value="historical_income">Facturas históricas</option><option value="quarterly_expense_import">Gastos trimestrales</option></select></div>';
        echo '</div><p><button class="vvd-btn" type="submit">Importar CSV</button></p></form></div></div>';
        echo '<div class="vvd-col-6"><div class="vvd-card"><h2>Importar histórico desde Excel</h2><p class="vvd-note">Carga un XLSX del programa anterior para completar la información de las facturas antiguas y meterlas en la tabla principal. Así las históricas y las nuevas encajan en la misma vista de Facturación.</p><form method="post" action="' . esc_url(admin_url('admin-post.php')) . '" enctype="multipart/form-data">';
        wp_nonce_field('vvd_import_legacy_xlsx');
        echo '<input type="hidden" name="action" value="vvd_import_legacy_xlsx">';
        echo '<div class="vvd-grid">';
        echo '<div class="vvd-col-12"><label class="vvd-label">Archivo XLSX</label><input class="vvd-input" type="file" name="legacy_xlsx" accept=".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" required></div>';
        echo '<div class="vvd-col-12"><label class="vvd-label">Modo</label><select class="vvd-select" name="legacy_mode"><option value="merge">Completar y actualizar si ya existe</option><option value="import_only_missing">Solo importar las que falten</option></select></div>';
        echo '</div><p><button class="vvd-btn" type="submit">Importar Excel histórico</button></p></form></div></div>';
        echo '<div class="vvd-col-6"><div class="vvd-card"><h2>Importar clientes desde Excel</h2><p class="vvd-note">Carga el XLSX de contactos del programa anterior para crear clientes nuevos o completar los existentes por NIF, email, ID externo o nombre. También intenta enlazar facturas históricas con la ficha del cliente correcta.</p><form method="post" action="' . esc_url(admin_url('admin-post.php')) . '" enctype="multipart/form-data">';
        wp_nonce_field('vvd_import_clients_xlsx');
        echo '<input type="hidden" name="action" value="vvd_import_clients_xlsx">';
        echo '<div class="vvd-grid">';
        echo '<div class="vvd-col-12"><label class="vvd-label">Archivo XLSX</label><input class="vvd-input" type="file" name="clients_xlsx" accept=".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" required></div>';
        echo '<div class="vvd-col-12"><label class="vvd-label">Modo</label><select class="vvd-select" name="clients_mode"><option value="merge">Crear y completar clientes existentes</option><option value="import_only_missing">Solo crear los que falten</option></select></div>';
        echo '</div><p><button class="vvd-btn" type="submit">Importar clientes</button></p></form></div></div>';
        global $wpdb;
        $legacy_invoice_count = (int) $wpdb->get_var("SELECT COUNT(*) FROM " . self::table('vvd_invoices') . " WHERE source_origin='legacy_import'");
        $legacy_line_count = (int) $wpdb->get_var("SELECT COUNT(*) FROM " . self::table('vvd_invoice_lines') . " l INNER JOIN " . self::table('vvd_invoices') . " i ON i.id=l.invoice_id WHERE i.source_origin='legacy_import'");
        $imported_clients_total = (int) $wpdb->get_var("SELECT COUNT(*) FROM " . self::table('vvd_clients') . " WHERE " . self::imported_clients_where_sql());
        $imported_clients_linked = (int) $wpdb->get_var("SELECT COUNT(*) FROM " . self::table('vvd_clients') . " c WHERE (" . self::imported_clients_where_sql() . ") AND EXISTS (SELECT 1 FROM " . self::table('vvd_invoices') . " i WHERE i.client_id=c.id)");
        $imported_clients_deletable = max(0, $imported_clients_total - $imported_clients_linked);

        echo '<div class="vvd-col-6"><div class="vvd-card"><h2>Eliminar facturas históricas subidas</h2><p class="vvd-note">Borra las facturas importadas desde el Excel histórico, sus líneas, sus eventos y los registros de importación. No toca las facturas nuevas del plugin.</p><p><strong>Facturas:</strong> ' . esc_html($legacy_invoice_count) . ' · <strong>Líneas:</strong> ' . esc_html($legacy_line_count) . '</p><form method="post" action="' . esc_url(admin_url('admin-post.php')) . '" onsubmit="return confirm(&quot;¿Eliminar todas las facturas históricas importadas?&quot;)">';
        wp_nonce_field('vvd_delete_legacy_imports');
        echo '<input type="hidden" name="action" value="vvd_delete_legacy_imports"><button class="vvd-btn danger" type="submit">Eliminar facturas históricas</button></form></div></div>';

        echo '<div class="vvd-col-6"><div class="vvd-card"><h2>Eliminar clientes subidos</h2><p class="vvd-note">Borra los clientes importados desde el Excel de contactos. Por seguridad, los que ya tengan documentos asociados no se borran.</p><p><strong>Importados:</strong> ' . esc_html($imported_clients_total) . ' · <strong>Se pueden borrar:</strong> ' . esc_html($imported_clients_deletable) . ' · <strong>Con documentos:</strong> ' . esc_html($imported_clients_linked) . '</p><form method="post" action="' . esc_url(admin_url('admin-post.php')) . '" onsubmit="return confirm(&quot;¿Eliminar los clientes importados que no tienen documentos asociados?&quot;)">';
        wp_nonce_field('vvd_delete_imported_clients');
        echo '<input type="hidden" name="action" value="vvd_delete_imported_clients"><button class="vvd-btn danger" type="submit">Eliminar clientes subidos</button></form></div></div>';
        echo '</div>';

        $imports = $wpdb->get_results('SELECT * FROM ' . self::imports_table() . ' ORDER BY imported_at DESC, id DESC LIMIT 20', ARRAY_A);
        echo '<div class="vvd-card"><h2>Últimas importaciones</h2><table class="vvd-table"><thead><tr><th>Fecha</th><th>Archivo</th><th>Tipo</th><th>Filas</th><th>Válidas</th><th>Errores</th></tr></thead><tbody>';
        foreach ($imports as $row) {
            echo '<tr><td>' . esc_html($row['imported_at']) . '</td><td>' . esc_html($row['filename']) . '</td><td>' . esc_html($row['source_type']) . '</td><td>' . esc_html($row['total_rows']) . '</td><td>' . esc_html($row['valid_rows']) . '</td><td>' . esc_html($row['error_rows']) . '</td></tr>';
        }
        if (!$imports) {
            echo '<tr><td colspan="6">No se ha importado nada todavía.</td></tr>';
        }
        echo '</tbody></table></div></div>';
    }

    public static function import_accounting_csv() {
        if (!current_user_can('manage_options')) { wp_die('No autorizado'); }
        check_admin_referer('vvd_import_accounting_csv');
        if (empty($_FILES['accounting_csv']['tmp_name'])) {
            wp_safe_redirect(admin_url('admin.php?page=vvd_importar_csv&vvd_import_status=error&vvd_message=' . rawurlencode('No se ha recibido ningún archivo.')));
            exit;
        }
        $file = $_FILES['accounting_csv'];
        $tmp = $file['tmp_name'];
        $filename = sanitize_file_name($file['name'] ?? 'import.csv');
        $hash = hash_file('sha256', $tmp);
        global $wpdb;
        $imports_table = self::imports_table();
        $entries_table = self::accounting_entries_table();
        $existing = (int) $wpdb->get_var($wpdb->prepare("SELECT id FROM {$imports_table} WHERE hash_file=%s LIMIT 1", $hash));
        if ($existing) {
            wp_safe_redirect(admin_url('admin.php?page=vvd_importar_csv&vvd_import_status=error&vvd_message=' . rawurlencode('Ese CSV ya fue importado antes.')));
            exit;
        }
        $parsed = self::parse_accounting_csv($tmp, $filename);
        $rows = $parsed['rows'];
        if (!$rows) {
            wp_safe_redirect(admin_url('admin.php?page=vvd_importar_csv&vvd_import_status=error&vvd_message=' . rawurlencode('El CSV no contiene filas importables.')));
            exit;
        }
        $source_type = sanitize_text_field($_POST['source_type'] ?? 'auto');
        if ($source_type === 'auto') {
            $source_type = $rows[0]['source_period_type'] ?? 'manual_import';
        }
        $wpdb->insert($imports_table, [
            'filename' => $filename,
            'source_type' => $source_type,
            'hash_file' => $hash,
            'total_rows' => count($rows),
            'valid_rows' => count($rows),
            'error_rows' => count($parsed['errors']),
            'meta_json' => wp_json_encode(['errors' => $parsed['errors']]),
            'imported_by' => get_current_user_id(),
            'imported_at' => current_time('mysql'),
        ]);
        $import_id = (int) $wpdb->insert_id;
        foreach ($rows as $row) {
            $row['import_id'] = $import_id;
            if ($source_type !== 'auto') {
                $row['source_period_type'] = $source_type;
            }
            $wpdb->insert($entries_table, $row);
            self::maybe_sync_client_from_entry($row);
        }
        self::log_event('accounting_csv_imported', 'Importado CSV contable: ' . $filename, null, null, ['import_id' => $import_id, 'rows' => count($rows), 'source_type' => $source_type]);
        wp_safe_redirect(admin_url('admin.php?page=vvd_importar_csv&vvd_import_status=ok&vvd_message=' . rawurlencode('CSV importado correctamente. Filas: ' . count($rows))));
        exit;
    }

}