<?php
if (!defined('ABSPATH')) {
    exit;
}

class VVD_Install {
    public static function activate() {
        global $wpdb;
        require_once ABSPATH . 'wp-admin/includes/upgrade.php';

        $charset = $wpdb->get_charset_collate();
        $clients = $wpdb->prefix . 'vvd_clients';
        $invoices = $wpdb->prefix . 'vvd_invoices';
        $lines = $wpdb->prefix . 'vvd_invoice_lines';
        $settings = $wpdb->prefix . 'vvd_settings';
        $events = $wpdb->prefix . 'vvd_events';
        $series = $wpdb->prefix . 'vvd_series';
        $imports = $wpdb->prefix . 'vvd_imports';
        $entries = $wpdb->prefix . 'vvd_accounting_entries';

        dbDelta("CREATE TABLE $clients (
            id BIGINT UNSIGNED NOT NULL AUTO_INCREMENT,
            name VARCHAR(255) NOT NULL,
            tax_id VARCHAR(80) NOT NULL,
            address TEXT NULL,
            email VARCHAR(255) NULL,
            phone VARCHAR(60) NULL,
            created_at DATETIME NOT NULL,
            updated_at DATETIME NULL,
            PRIMARY KEY (id),
            KEY tax_id (tax_id)
        ) $charset;");

        dbDelta("CREATE TABLE $invoices (
            id BIGINT UNSIGNED NOT NULL AUTO_INCREMENT,
            series VARCHAR(20) NOT NULL,
            series_label VARCHAR(60) NULL,
            number INT NOT NULL,
            code VARCHAR(60) NULL,
            invoice_type VARCHAR(30) NOT NULL DEFAULT 'standard',
            issue_date DATE NOT NULL,
            due_date DATE NULL,
            client_id BIGINT UNSIGNED NOT NULL,
            original_invoice_id BIGINT UNSIGNED NULL,
            notes TEXT NULL,
            taxable_base DECIMAL(12,2) NOT NULL DEFAULT 0,
            tax_rate DECIMAL(7,2) NOT NULL DEFAULT 21.00,
            tax_amount DECIMAL(12,2) NOT NULL DEFAULT 0,
            irpf_rate DECIMAL(7,2) NOT NULL DEFAULT 0,
            irpf_amount DECIMAL(12,2) NOT NULL DEFAULT 0,
            total_amount DECIMAL(12,2) NOT NULL DEFAULT 0,
            status VARCHAR(50) NOT NULL DEFAULT 'issued',
            pdf_path TEXT NULL,
            sent_at DATETIME NULL,
            paid_at DATETIME NULL,
            verifactu_status VARCHAR(50) NOT NULL DEFAULT 'draft-ready',
            verifactu_external_id VARCHAR(120) NULL,
            hash_prev VARCHAR(128) NULL,
            hash_current VARCHAR(128) NULL,
            created_at DATETIME NOT NULL,
            updated_at DATETIME NULL,
            PRIMARY KEY (id),
            UNIQUE KEY code (code),
            KEY series_number (series, number),
            KEY client_id (client_id),
            KEY original_invoice_id (original_invoice_id),
            KEY verifactu_status (verifactu_status),
            KEY invoice_type (invoice_type)
        ) $charset;");

        dbDelta("CREATE TABLE $lines (
            id BIGINT UNSIGNED NOT NULL AUTO_INCREMENT,
            invoice_id BIGINT UNSIGNED NOT NULL,
            description VARCHAR(255) NOT NULL,
            quantity DECIMAL(12,2) NOT NULL DEFAULT 1.00,
            unit_price DECIMAL(12,2) NOT NULL DEFAULT 0.00,
            line_total DECIMAL(12,2) NOT NULL DEFAULT 0.00,
            PRIMARY KEY (id),
            KEY invoice_id (invoice_id)
        ) $charset;");

        dbDelta("CREATE TABLE $settings (
            id BIGINT UNSIGNED NOT NULL AUTO_INCREMENT,
            issuer_name VARCHAR(255) NOT NULL DEFAULT '',
            issuer_tax_id VARCHAR(80) NOT NULL DEFAULT '',
            issuer_address TEXT NULL,
            issuer_email VARCHAR(255) NOT NULL DEFAULT '',
            issuer_phone VARCHAR(60) NOT NULL DEFAULT '',
            issuer_iban VARCHAR(80) NOT NULL DEFAULT '',
            default_series VARCHAR(20) NOT NULL DEFAULT 'A',
            rectificative_series VARCHAR(20) NOT NULL DEFAULT 'RT',
            quote_series VARCHAR(20) NOT NULL DEFAULT 'P',
            series_padding INT NOT NULL DEFAULT 6,
            default_tax_rate DECIMAL(7,2) NOT NULL DEFAULT 21.00,
            default_irpf_rate DECIMAL(7,2) NOT NULL DEFAULT 0.00,
            logo_url TEXT NULL,
            background_image_url TEXT NULL,
            brand_primary VARCHAR(20) NOT NULL DEFAULT '#6997c1',
            brand_secondary VARCHAR(20) NOT NULL DEFAULT '#8d8e8e',
            brand_background VARCHAR(20) NOT NULL DEFAULT '#FFFFFF',
            created_at DATETIME NOT NULL,
            updated_at DATETIME NOT NULL,
            PRIMARY KEY (id)
        ) $charset;");

        dbDelta("CREATE TABLE $events (
            id BIGINT UNSIGNED NOT NULL AUTO_INCREMENT,
            invoice_id BIGINT UNSIGNED NULL,
            client_id BIGINT UNSIGNED NULL,
            event_type VARCHAR(80) NOT NULL,
            event_message TEXT NULL,
            context LONGTEXT NULL,
            user_id BIGINT UNSIGNED NULL,
            created_at DATETIME NOT NULL,
            PRIMARY KEY (id),
            KEY invoice_id (invoice_id),
            KEY client_id (client_id),
            KEY event_type (event_type)
        ) $charset;");

        dbDelta("CREATE TABLE $series (
            id BIGINT UNSIGNED NOT NULL AUTO_INCREMENT,
            series_key VARCHAR(20) NOT NULL,
            label_text VARCHAR(100) NULL,
            prefix VARCHAR(20) NOT NULL,
            padding INT NOT NULL DEFAULT 6,
            current_number INT NOT NULL DEFAULT 0,
            invoice_type VARCHAR(30) NOT NULL DEFAULT 'standard',
            is_active TINYINT(1) NOT NULL DEFAULT 1,
            created_at DATETIME NOT NULL,
            updated_at DATETIME NULL,
            PRIMARY KEY (id),
            UNIQUE KEY series_key (series_key)
        ) $charset;");


        dbDelta("CREATE TABLE $imports (
            id BIGINT UNSIGNED NOT NULL AUTO_INCREMENT,
            filename VARCHAR(255) NOT NULL,
            source_type VARCHAR(80) NOT NULL DEFAULT 'manual',
            hash_file VARCHAR(64) NOT NULL,
            total_rows INT NOT NULL DEFAULT 0,
            valid_rows INT NOT NULL DEFAULT 0,
            error_rows INT NOT NULL DEFAULT 0,
            meta_json LONGTEXT NULL,
            imported_by BIGINT UNSIGNED NULL,
            imported_at DATETIME NOT NULL,
            PRIMARY KEY (id),
            UNIQUE KEY hash_file (hash_file),
            KEY source_type (source_type)
        ) $charset;");

        dbDelta("CREATE TABLE $entries (
            id BIGINT UNSIGNED NOT NULL AUTO_INCREMENT,
            import_id BIGINT UNSIGNED NOT NULL,
            entry_date DATE NOT NULL,
            nif VARCHAR(80) NULL,
            account_code VARCHAR(40) NULL,
            description TEXT NULL,
            journal_entry VARCHAR(80) NULL,
            invoice_number_raw VARCHAR(120) NULL,
            invoice_number_normalized VARCHAR(120) NULL,
            base_amount DECIMAL(12,2) NOT NULL DEFAULT 0.00,
            tax_rate DECIMAL(7,2) NOT NULL DEFAULT 0.00,
            tax_amount DECIMAL(12,2) NOT NULL DEFAULT 0.00,
            irpf_rate DECIMAL(7,2) NOT NULL DEFAULT 0.00,
            irpf_amount DECIMAL(12,2) NOT NULL DEFAULT 0.00,
            total_amount DECIMAL(12,2) NOT NULL DEFAULT 0.00,
            entry_group VARCHAR(20) NOT NULL DEFAULT 'other',
            document_type VARCHAR(30) NOT NULL DEFAULT 'unknown',
            business_line VARCHAR(30) NOT NULL DEFAULT 'other',
            series_code VARCHAR(30) NOT NULL DEFAULT 'NONE',
            source_file VARCHAR(255) NULL,
            source_period_type VARCHAR(40) NULL,
            raw_row_json LONGTEXT NULL,
            created_at DATETIME NOT NULL,
            PRIMARY KEY (id),
            KEY import_id (import_id),
            KEY entry_date (entry_date),
            KEY entry_group (entry_group),
            KEY business_line (business_line),
            KEY document_type (document_type),
            KEY invoice_number_normalized (invoice_number_normalized),
            KEY nif (nif),
            KEY account_code (account_code)
        ) $charset;");



        self::maybe_add_column($clients, 'external_id', "ALTER TABLE $clients ADD COLUMN external_id VARCHAR(120) NULL AFTER tax_id");
        self::maybe_add_column($clients, 'mobile', "ALTER TABLE $clients ADD COLUMN mobile VARCHAR(60) NULL AFTER phone");
        self::maybe_add_column($clients, 'city', "ALTER TABLE $clients ADD COLUMN city VARCHAR(120) NULL AFTER address");
        self::maybe_add_column($clients, 'state', "ALTER TABLE $clients ADD COLUMN state VARCHAR(120) NULL AFTER city");
        self::maybe_add_column($clients, 'postcode', "ALTER TABLE $clients ADD COLUMN postcode VARCHAR(30) NULL AFTER state");
        self::maybe_add_column($clients, 'country', "ALTER TABLE $clients ADD COLUMN country VARCHAR(120) NULL AFTER postcode");
        self::maybe_add_column($clients, 'display_name', "ALTER TABLE $clients ADD COLUMN display_name VARCHAR(255) NULL AFTER name");
        self::maybe_add_column($clients, 'company_name', "ALTER TABLE $clients ADD COLUMN company_name VARCHAR(255) NULL AFTER display_name");
        self::maybe_add_column($clients, 'salutation', "ALTER TABLE $clients ADD COLUMN salutation VARCHAR(60) NULL AFTER company_name");
        self::maybe_add_column($clients, 'first_name', "ALTER TABLE $clients ADD COLUMN first_name VARCHAR(120) NULL AFTER salutation");
        self::maybe_add_column($clients, 'last_name', "ALTER TABLE $clients ADD COLUMN last_name VARCHAR(120) NULL AFTER first_name");
        self::maybe_add_column($clients, 'website', "ALTER TABLE $clients ADD COLUMN website VARCHAR(255) NULL AFTER email");
        self::maybe_add_column($clients, 'client_status', "ALTER TABLE $clients ADD COLUMN client_status VARCHAR(60) NULL AFTER website");
        self::maybe_add_column($clients, 'currency_code', "ALTER TABLE $clients ADD COLUMN currency_code VARCHAR(12) NULL AFTER client_status");
        self::maybe_add_column($clients, 'notes', "ALTER TABLE $clients ADD COLUMN notes TEXT NULL AFTER currency_code");
        self::maybe_add_column($clients, 'billing_attention', "ALTER TABLE $clients ADD COLUMN billing_attention VARCHAR(255) NULL AFTER address");
        self::maybe_add_column($clients, 'billing_street2', "ALTER TABLE $clients ADD COLUMN billing_street2 VARCHAR(255) NULL AFTER billing_attention");
        self::maybe_add_column($clients, 'shipping_attention', "ALTER TABLE $clients ADD COLUMN shipping_attention VARCHAR(255) NULL AFTER country");
        self::maybe_add_column($clients, 'shipping_address', "ALTER TABLE $clients ADD COLUMN shipping_address TEXT NULL AFTER shipping_attention");
        self::maybe_add_column($clients, 'shipping_street2', "ALTER TABLE $clients ADD COLUMN shipping_street2 VARCHAR(255) NULL AFTER shipping_address");
        self::maybe_add_column($clients, 'shipping_city', "ALTER TABLE $clients ADD COLUMN shipping_city VARCHAR(120) NULL AFTER shipping_street2");
        self::maybe_add_column($clients, 'shipping_state', "ALTER TABLE $clients ADD COLUMN shipping_state VARCHAR(120) NULL AFTER shipping_city");
        self::maybe_add_column($clients, 'shipping_postcode', "ALTER TABLE $clients ADD COLUMN shipping_postcode VARCHAR(30) NULL AFTER shipping_state");
        self::maybe_add_column($clients, 'shipping_country', "ALTER TABLE $clients ADD COLUMN shipping_country VARCHAR(120) NULL AFTER shipping_postcode");
        self::maybe_add_column($clients, 'payment_terms', "ALTER TABLE $clients ADD COLUMN payment_terms VARCHAR(120) NULL AFTER shipping_country");
        self::maybe_add_column($clients, 'payment_terms_label', "ALTER TABLE $clients ADD COLUMN payment_terms_label VARCHAR(255) NULL AFTER payment_terms");
        self::maybe_add_column($clients, 'contact_name', "ALTER TABLE $clients ADD COLUMN contact_name VARCHAR(255) NULL AFTER payment_terms_label");
        self::maybe_add_column($clients, 'contact_type', "ALTER TABLE $clients ADD COLUMN contact_type VARCHAR(80) NULL AFTER contact_name");
        self::maybe_add_column($clients, 'designation', "ALTER TABLE $clients ADD COLUMN designation VARCHAR(120) NULL AFTER contact_type");
        self::maybe_add_column($clients, 'department', "ALTER TABLE $clients ADD COLUMN department VARCHAR(120) NULL AFTER designation");
        self::maybe_add_column($clients, 'taxable', "ALTER TABLE $clients ADD COLUMN taxable VARCHAR(20) NULL AFTER department");
        self::maybe_add_column($clients, 'tax_name', "ALTER TABLE $clients ADD COLUMN tax_name VARCHAR(120) NULL AFTER taxable");
        self::maybe_add_column($clients, 'tax_percentage', "ALTER TABLE $clients ADD COLUMN tax_percentage VARCHAR(40) NULL AFTER tax_name");
        self::maybe_add_column($clients, 'exemption_reason', "ALTER TABLE $clients ADD COLUMN exemption_reason TEXT NULL AFTER tax_percentage");
        self::maybe_add_column($clients, 'contact_id_legacy', "ALTER TABLE $clients ADD COLUMN contact_id_legacy VARCHAR(120) NULL AFTER exemption_reason");
        self::maybe_add_column($clients, 'primary_contact_id', "ALTER TABLE $clients ADD COLUMN primary_contact_id VARCHAR(120) NULL AFTER contact_id_legacy");
        self::maybe_add_column($clients, 'contacts_json', "ALTER TABLE $clients ADD COLUMN contacts_json LONGTEXT NULL AFTER primary_contact_id");
        self::maybe_add_column($clients, 'raw_import_json', "ALTER TABLE $clients ADD COLUMN raw_import_json LONGTEXT NULL AFTER contacts_json");

        self::maybe_add_column($invoices, 'source_origin', "ALTER TABLE $invoices ADD COLUMN source_origin VARCHAR(30) NOT NULL DEFAULT 'plugin_generated' AFTER invoice_type");
        self::maybe_add_column($invoices, 'external_invoice_id', "ALTER TABLE $invoices ADD COLUMN external_invoice_id VARCHAR(120) NULL AFTER source_origin");
        self::maybe_add_column($invoices, 'external_invoice_number', "ALTER TABLE $invoices ADD COLUMN external_invoice_number VARCHAR(120) NULL AFTER external_invoice_id");
        self::maybe_add_column($invoices, 'client_external_id', "ALTER TABLE $invoices ADD COLUMN client_external_id VARCHAR(120) NULL AFTER client_id");
        self::maybe_add_column($invoices, 'legacy_status', "ALTER TABLE $invoices ADD COLUMN legacy_status VARCHAR(80) NULL AFTER status");
        self::maybe_add_column($invoices, 'billing_name', "ALTER TABLE $invoices ADD COLUMN billing_name VARCHAR(255) NULL AFTER client_external_id");
        self::maybe_add_column($invoices, 'billing_address', "ALTER TABLE $invoices ADD COLUMN billing_address TEXT NULL AFTER billing_name");
        self::maybe_add_column($invoices, 'billing_city', "ALTER TABLE $invoices ADD COLUMN billing_city VARCHAR(120) NULL AFTER billing_address");
        self::maybe_add_column($invoices, 'billing_state', "ALTER TABLE $invoices ADD COLUMN billing_state VARCHAR(120) NULL AFTER billing_city");
        self::maybe_add_column($invoices, 'billing_postcode', "ALTER TABLE $invoices ADD COLUMN billing_postcode VARCHAR(30) NULL AFTER billing_state");
        self::maybe_add_column($invoices, 'billing_country', "ALTER TABLE $invoices ADD COLUMN billing_country VARCHAR(120) NULL AFTER billing_postcode");
        self::maybe_add_column($invoices, 'client_email', "ALTER TABLE $invoices ADD COLUMN client_email VARCHAR(255) NULL AFTER billing_country");
        self::maybe_add_column($invoices, 'client_phone', "ALTER TABLE $invoices ADD COLUMN client_phone VARCHAR(60) NULL AFTER client_email");
        self::maybe_add_column($invoices, 'client_mobile', "ALTER TABLE $invoices ADD COLUMN client_mobile VARCHAR(60) NULL AFTER client_phone");
        self::maybe_add_column($invoices, 'currency_code', "ALTER TABLE $invoices ADD COLUMN currency_code VARCHAR(12) NULL AFTER client_mobile");
        self::maybe_add_column($invoices, 'payment_terms', "ALTER TABLE $invoices ADD COLUMN payment_terms VARCHAR(120) NULL AFTER currency_code");
        self::maybe_add_column($invoices, 'payment_terms_label', "ALTER TABLE $invoices ADD COLUMN payment_terms_label VARCHAR(255) NULL AFTER payment_terms");
        self::maybe_add_column($invoices, 'expected_payment_date', "ALTER TABLE $invoices ADD COLUMN expected_payment_date DATE NULL AFTER payment_terms_label");
        self::maybe_add_column($invoices, 'last_payment_date', "ALTER TABLE $invoices ADD COLUMN last_payment_date DATE NULL AFTER expected_payment_date");
        self::maybe_add_column($invoices, 'terms_conditions', "ALTER TABLE $invoices ADD COLUMN terms_conditions TEXT NULL AFTER notes");
        self::maybe_add_column($invoices, 'estimate_number', "ALTER TABLE $invoices ADD COLUMN estimate_number VARCHAR(120) NULL AFTER terms_conditions");
        self::maybe_add_column($invoices, 'subtotal', "ALTER TABLE $invoices ADD COLUMN subtotal DECIMAL(12,2) NOT NULL DEFAULT 0.00 AFTER estimate_number");
        self::maybe_add_column($invoices, 'discount_total', "ALTER TABLE $invoices ADD COLUMN discount_total DECIMAL(12,2) NOT NULL DEFAULT 0.00 AFTER subtotal");
        self::maybe_add_column($invoices, 'withholding_total', "ALTER TABLE $invoices ADD COLUMN withholding_total DECIMAL(12,2) NOT NULL DEFAULT 0.00 AFTER discount_total");
        self::maybe_add_column($invoices, 'balance_due', "ALTER TABLE $invoices ADD COLUMN balance_due DECIMAL(12,2) NOT NULL DEFAULT 0.00 AFTER withholding_total");
        self::maybe_add_column($invoices, 'import_hash', "ALTER TABLE $invoices ADD COLUMN import_hash VARCHAR(64) NULL AFTER balance_due");
        self::maybe_add_column($invoices, 'paid_at', "ALTER TABLE $invoices ADD COLUMN paid_at DATETIME NULL AFTER sent_at");

        self::maybe_add_column($lines, 'item_name', "ALTER TABLE $lines ADD COLUMN item_name VARCHAR(255) NULL AFTER invoice_id");
        self::maybe_add_column($lines, 'item_desc', "ALTER TABLE $lines ADD COLUMN item_desc TEXT NULL AFTER item_name");
        self::maybe_add_column($lines, 'usage_unit', "ALTER TABLE $lines ADD COLUMN usage_unit VARCHAR(60) NULL AFTER item_desc");
        self::maybe_add_column($lines, 'discount_amount', "ALTER TABLE $lines ADD COLUMN discount_amount DECIMAL(12,2) NOT NULL DEFAULT 0.00 AFTER quantity");
        self::maybe_add_column($lines, 'sort_order', "ALTER TABLE $lines ADD COLUMN sort_order INT NOT NULL DEFAULT 0 AFTER line_total");

        $exists = (int) $wpdb->get_var("SELECT COUNT(*) FROM $settings");
        $now = current_time('mysql');
        if (!$exists) {
            $wpdb->insert($settings, [
                'issuer_name' => get_bloginfo('name'),
                'default_series' => 'A',
                'rectificative_series' => 'RT',
                'quote_series' => 'P',
                'series_padding' => 6,
                'default_tax_rate' => 21.00,
                'default_irpf_rate' => 15.00,
                'logo_url' => 'https://vexeldot.es/wp-content/uploads/cropped-Logo-VexelDot.png',
                'background_image_url' => 'https://vexeldot.es/wp-content/uploads/fondo-factura.jpg',
                'brand_primary' => '#6997c1',
                'brand_secondary' => '#8d8e8e',
                'brand_background' => '#FFFFFF',
                'created_at' => $now,
                'updated_at' => $now,
            ]);
        } else {
            $wpdb->query("UPDATE $settings SET quote_series='P' WHERE quote_series IS NULL OR quote_series='' ");
        }

        self::ensure_series('A', 'Serie estándar', 'A', 6, 'standard');
        self::ensure_series('RT', 'Serie rectificativa', 'RT', 6, 'rectificativa');
        self::ensure_series('P', 'Serie presupuestos', 'P', 6, 'quote');

        update_option('vvd_plugin_version', VVD_VERSION);
    }

    public static function maybe_upgrade() {
        $installed = get_option('vvd_plugin_version');
        if (!$installed || version_compare((string) $installed, VVD_VERSION, '<')) {
            self::activate();
        }
    }


    protected static function maybe_add_column($table, $column, $sql) {
        global $wpdb;
        $exists = $wpdb->get_var($wpdb->prepare("SHOW COLUMNS FROM {$table} LIKE %s", $column));
        if (!$exists) {
            $wpdb->query($sql);
        }
    }

    protected static function ensure_series($series_key, $label, $prefix, $padding, $type) {
        global $wpdb;
        $table = $wpdb->prefix . 'vvd_series';
        $exists = $wpdb->get_var($wpdb->prepare("SELECT id FROM $table WHERE series_key=%s", $series_key));
        if (!$exists) {
            $wpdb->insert($table, [
                'series_key' => $series_key,
                'label_text' => $label,
                'prefix' => $prefix,
                'padding' => $padding,
                'current_number' => 0,
                'invoice_type' => $type,
                'is_active' => 1,
                'created_at' => current_time('mysql'),
            ]);
        }
    }
}
