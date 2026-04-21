<?php
/**
 * Plugin Name: VexelDot Facturación
 * Description: Gestión simple de clientes y facturas con rectificativas, duplicado, PDF visual y envío por correo. Base propia, no sustituye el cumplimiento legal completo de VERI*FACTU.
 * Version: 0.8.8
 * Author: OpenAI para VexelDot
 * Requires at least: 6.0
 * Requires PHP: 7.4
 */

if (!defined('ABSPATH')) {
    exit;
}

define('VVD_PLUGIN_FILE', __FILE__);
define('VVD_PLUGIN_DIR', plugin_dir_path(__FILE__));
define('VVD_PLUGIN_URL', plugin_dir_url(__FILE__));
define('VVD_VERSION', '0.8.8');

require_once VVD_PLUGIN_DIR . 'includes/class-vvd-install.php';
require_once VVD_PLUGIN_DIR . 'includes/class-vvd-pdf.php';
require_once VVD_PLUGIN_DIR . 'includes/class-vvd-admin.php';

register_activation_hook(__FILE__, ['VVD_Install', 'activate']);
add_action('plugins_loaded', ['VVD_Install', 'maybe_upgrade']);
add_action('plugins_loaded', ['VVD_Admin', 'init']);
