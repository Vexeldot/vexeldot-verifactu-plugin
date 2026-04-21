<?php
if (!defined('ABSPATH')) {
    exit;
}

class VVD_PDF {
    protected $objects = [];

    protected function add_object($content) {
        $this->objects[] = $content;
        return count($this->objects);
    }

    protected function color_to_rgb($hex) {
        $hex = ltrim((string) $hex, '#');
        if (strlen($hex) === 3) {
            $hex = $hex[0].$hex[0].$hex[1].$hex[1].$hex[2].$hex[2];
        }
        if (strlen($hex) !== 6) {
            $hex = '000000';
        }
        return [hexdec(substr($hex, 0, 2)) / 255, hexdec(substr($hex, 2, 2)) / 255, hexdec(substr($hex, 4, 2)) / 255];
    }

    protected function esc($text) {
        $map = [
            '€' => 'EUR', 'á' => 'a', 'é' => 'e', 'í' => 'i', 'ó' => 'o', 'ú' => 'u',
            'Á' => 'A', 'É' => 'E', 'Í' => 'I', 'Ó' => 'O', 'Ú' => 'U', 'ñ' => 'n', 'Ñ' => 'N',
            'ü' => 'u', 'Ü' => 'U', 'ç' => 'c', 'Ç' => 'C',
            'º' => 'o', 'ª' => 'a', '–' => '-', '—' => '-',
        ];
        $text = strtr((string) $text, $map);
        $text = preg_replace('/[^\x20-\x7E]/', '', $text);
        return str_replace(['\\', '(', ')'], ['\\\\', '\\(', '\\)'], $text);
    }

    protected function fmt_money($amount) {
        return number_format((float) $amount, 2, ',', '.') . ' EUR';
    }

    protected function wrap_text($text, $max_width, $font_size = 10) {
        $text = trim(preg_replace('/\s+/', ' ', (string) $text));
        if ($text === '') {
            return ['-'];
        }
        $avg = max(1, $font_size * 0.48);
        $max_chars = max(8, (int) floor($max_width / $avg));
        return explode("\n", wordwrap($text, $max_chars, "\n", true));
    }

    protected function text_width($text, $font_size = 10) {
        return strlen((string) $text) * max(1, $font_size * 0.48);
    }

    protected function text_cmd($x, $y, $text, $font = 'F1', $size = 10, $rgb = [0, 0, 0]) {
        return sprintf(
            "BT /%s %s Tf %.3F %.3F %.3F rg %.2F %.2F Td (%s) Tj ET\n",
            $font,
            $size,
            $rgb[0],
            $rgb[1],
            $rgb[2],
            $x,
            $y,
            $this->esc($text)
        );
    }

    protected function text_right_cmd($x_right, $y, $text, $font = 'F1', $size = 10, $rgb = [0, 0, 0]) {
        $x = $x_right - $this->text_width($text, $size);
        return $this->text_cmd($x, $y, $text, $font, $size, $rgb);
    }

    protected function text_center_cmd($x_center, $y, $text, $font = 'F1', $size = 10, $rgb = [0, 0, 0]) {
        $x = $x_center - ($this->text_width($text, $size) / 2);
        return $this->text_cmd($x, $y, $text, $font, $size, $rgb);
    }

    protected function jpeg_from_url($url) {
        if (!$url) {
            return null;
        }
        $response = wp_remote_get($url, ['timeout' => 20]);
        if (is_wp_error($response)) {
            return null;
        }
        $body = wp_remote_retrieve_body($response);
        if (!$body) {
            return null;
        }
        $img = @imagecreatefromstring($body);
        if (!$img) {
            return null;
        }
        ob_start();
        imagejpeg($img, null, 90);
        $jpg = ob_get_clean();
        imagedestroy($img);
        if (!$jpg) {
            return null;
        }
        $info = @getimagesizefromstring($jpg);
        if (!$info) {
            return null;
        }
        return [
            'data' => $jpg,
            'width' => (int) $info[0],
            'height' => (int) $info[1],
        ];
    }

    public function render($invoice, $lines, $client, $settings) {
        $this->objects = [];
        $width = 595.28;
        $height = 841.89;
        $margin = 42;
        $content_width = $width - ($margin * 2);

        $primary = $this->color_to_rgb($settings['brand_primary'] ?: '#6997c1');
        $secondary = $this->color_to_rgb($settings['brand_secondary'] ?: '#8d8e8e');
        $white = $this->color_to_rgb('#ffffff');
        $text = $this->color_to_rgb('#2f3a45');
        $border = $this->color_to_rgb('#d9e1e8');
        $soft = $this->color_to_rgb('#f5f8fb');
        $muted = $this->color_to_rgb('#5f6c77');

        $stream = "q\n";
        $stream .= sprintf("%.3F %.3F %.3F rg 0 0 %.2F %.2F re f\n", $white[0], $white[1], $white[2], $width, $height);
        $stream .= sprintf("%.3F %.3F %.3F rg 0 %.2F %.2F 16 re f\n", $primary[0], $primary[1], $primary[2], $height - 16, $width);

        $image_info = $this->jpeg_from_url($settings['logo_url'] ?? '');
        $image_obj = 0;
        if ($image_info) {
            $image_dict = "<< /Type /XObject /Subtype /Image /Width {$image_info['width']} /Height {$image_info['height']} /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /DCTDecode /Length " . strlen($image_info['data']) . " >>\nstream\n" . $image_info['data'] . "\nendstream";
            $image_obj = $this->add_object($image_dict);
            $logo_w = 240;
            $logo_h = max(22, $logo_w * ($image_info['height'] / max(1, $image_info['width'])));
            $logo_y = $height - 112;
            $stream .= "q\n";
            $stream .= sprintf("%.2F 0 0 %.2F %.2F %.2F cm\n", $logo_w, $logo_h, $margin, $logo_y);
            $stream .= "/Im1 Do\nQ\n";
        } else {
            $stream .= $this->text_cmd($margin, $height - 72, $settings['issuer_name'] ?: 'VexelDot', 'F2', 18, $text);
        }

        $title = 'FACTURA';
        if (($invoice['invoice_type'] ?? '') === 'rectificativa') {
            $title = 'FACTURA RECTIFICATIVA';
        } elseif (($invoice['invoice_type'] ?? '') === 'quote') {
            $title = 'PRESUPUESTO';
        }

        $code = $invoice['code'] ?: ($invoice['series'] . '-' . str_pad((int) $invoice['number'], (int) ($settings['series_padding'] ?: 6), '0', STR_PAD_LEFT));
        $status = ucfirst((string) ($invoice['status'] ?: '-'));

        $doc_w = 186;
        $doc_h = 118;
        $doc_x = $width - $margin - $doc_w;
        $doc_y = $height - 168;
        $stream .= sprintf("1 1 1 rg %.2F %.2F %.2F %.2F re f\n", $doc_x, $doc_y, $doc_w, $doc_h);
        $stream .= sprintf("%.3F %.3F %.3F RG %.2F %.2F %.2F %.2F re S\n", $border[0], $border[1], $border[2], $doc_x, $doc_y, $doc_w, $doc_h);
        $stream .= $this->text_cmd($doc_x + 16, $doc_y + $doc_h - 18, 'DOCUMENTO', 'F2', 8, $secondary);
        $stream .= $this->text_cmd($doc_x + 16, $doc_y + $doc_h - 46, $title, 'F2', 18, $primary);

        $meta_rows = [
            ['Numero', $code],
            ['Fecha', $invoice['issue_date'] ?: '-'],
            ['Vencimiento', $invoice['due_date'] ?: '-'],
            ['Estado', $status],
        ];
        $my = $doc_y + $doc_h - 62;
        foreach ($meta_rows as $i => $row) {
            $y = $my - ($i * 14);
            $stream .= sprintf("%.3F %.3F %.3F RG %.2F %.2F m %.2F %.2F l S\n", $border[0], $border[1], $border[2], $doc_x + 16, $y - 4, $doc_x + $doc_w - 16, $y - 4);
            $stream .= $this->text_cmd($doc_x + 16, $y, $row[0], 'F1', 8, $secondary);
            $stream .= $this->text_right_cmd($doc_x + $doc_w - 16, $y, $row[1], 'F2', 8, $text);
        }

        $cards_y = $height - 314;
        $card_h = 116;
        $gap = 18;
        $card_w = ($content_width - $gap) / 2;
        $cards = [
            [$margin, 'EMISOR', [
                $settings['issuer_name'] ?: '-',
                'NIF: ' . ($settings['issuer_tax_id'] ?: '-'),
                $settings['issuer_address'] ?: '-',
                !empty($settings['issuer_email']) ? $settings['issuer_email'] : '',
                !empty($settings['issuer_phone']) ? $settings['issuer_phone'] : '',
            ]],
            [$margin + $card_w + $gap, 'CLIENTE', [
                $client['name'] ?: '-',
                'NIF: ' . ($client['tax_id'] ?: '-'),
                $client['address'] ?: '-',
                !empty($client['email']) ? $client['email'] : '',
                !empty($client['phone']) ? $client['phone'] : '',
            ]],
        ];
        foreach ($cards as $card) {
            [$x, $label, $rows] = $card;
            $stream .= sprintf("1 1 1 rg %.2F %.2F %.2F %.2F re f\n", $x, $cards_y, $card_w, $card_h);
            $stream .= sprintf("%.3F %.3F %.3F RG %.2F %.2F %.2F %.2F re S\n", $border[0], $border[1], $border[2], $x, $cards_y, $card_w, $card_h);
            $stream .= $this->text_cmd($x + 14, $cards_y + $card_h - 18, $label, 'F2', 9, $primary);
            $cy = $cards_y + $card_h - 36;
            foreach ($rows as $idx => $row) {
                if ($row === '') {
                    continue;
                }
                $wrapped = array_slice($this->wrap_text($row, $card_w - 28, $idx === 0 ? 10 : 9), 0, $idx === 0 ? 1 : 2);
                foreach ($wrapped as $line) {
                    if ($cy < ($cards_y + 12)) {
                        break 2;
                    }
                    $stream .= $this->text_cmd($x + 14, $cy, $line, $idx === 0 ? 'F2' : 'F1', $idx === 0 ? 10 : 9, $text);
                    $cy -= 11;
                }
            }
        }

        $table_top = $height - 396;
        $header_h = 26;
        $desc_w = 225;
        $qty_w = 50;
        $price_w = 78;
        $tax_w = 60;
        $total_w = $content_width - $desc_w - $qty_w - $price_w - $tax_w;
        $x_desc = $margin;
        $x_qty = $x_desc + $desc_w;
        $x_price = $x_qty + $qty_w;
        $x_tax = $x_price + $price_w;
        $x_total = $x_tax + $tax_w;

        $stream .= sprintf("%.3F %.3F %.3F rg %.2F %.2F %.2F %.2F re f\n", $primary[0], $primary[1], $primary[2], $margin, $table_top, $content_width, $header_h);
        $stream .= sprintf("%.3F %.3F %.3F RG %.2F %.2F %.2F %.2F re S\n", $border[0], $border[1], $border[2], $margin, $table_top, $content_width, $header_h);
        $stream .= $this->text_cmd($x_desc + 10, $table_top + 9, 'Concepto', 'F2', 8, $white);
        $stream .= $this->text_center_cmd($x_qty + ($qty_w / 2), $table_top + 9, 'Cant.', 'F2', 8, $white);
        $stream .= $this->text_right_cmd($x_price + $price_w - 10, $table_top + 9, 'Precio', 'F2', 8, $white);
        $stream .= $this->text_center_cmd($x_tax + ($tax_w / 2), $table_top + 9, 'Impuesto', 'F2', 8, $white);
        $stream .= $this->text_right_cmd($x_total + $total_w - 10, $table_top + 9, 'Importe', 'F2', 8, $white);

        $y = $table_top;
        foreach ($lines as $index => $line) {
            $desc = trim((string) ($line['description'] ?? ''));
            $desc_lines = array_slice($this->wrap_text($desc, $desc_w - 18, 9), 0, 4);
            $row_h = max(34, 14 + (count($desc_lines) * 11));
            if (($y - $row_h) < 216) {
                break;
            }
            $row_y = $y - $row_h;
            if ($index % 2 === 0) {
                $stream .= sprintf("1 1 1 rg %.2F %.2F %.2F %.2F re f\n", $margin, $row_y, $content_width, $row_h);
            } else {
                $stream .= sprintf("%.3F %.3F %.3F rg %.2F %.2F %.2F %.2F re f\n", $soft[0], $soft[1], $soft[2], $margin, $row_y, $content_width, $row_h);
            }
            $stream .= sprintf("%.3F %.3F %.3F RG %.2F %.2F %.2F %.2F re S\n", $border[0], $border[1], $border[2], $margin, $row_y, $content_width, $row_h);

            $ty = $y - 14;
            foreach ($desc_lines as $n => $dl) {
                $stream .= $this->text_cmd($x_desc + 10, $ty, $dl, $n === 0 ? 'F2' : 'F1', 9, $text);
                $ty -= 11;
            }

            $tax_rate = (($invoice['invoice_type'] ?? '') === 'quote') ? '0%' : (number_format((float) ($invoice['tax_rate'] ?? 0), 0, ',', '.') . '%');
            $qty_text = rtrim(rtrim(number_format((float) ($line['quantity'] ?? 0), 2, ',', '.'), '0'), ',');
            $stream .= $this->text_center_cmd($x_qty + ($qty_w / 2), $y - 14, $qty_text, 'F1', 9, $text);
            $stream .= $this->text_right_cmd($x_price + $price_w - 10, $y - 14, $this->fmt_money($line['unit_price'] ?? 0), 'F1', 9, $text);
            $stream .= $this->text_center_cmd($x_tax + ($tax_w / 2), $y - 14, $tax_rate, 'F1', 9, $text);
            $stream .= $this->text_right_cmd($x_total + $total_w - 10, $y - 14, $this->fmt_money($line['line_total'] ?? 0), 'F1', 9, $text);

            $y = $row_y;
        }

        $bottom_top = $y - 24;
        $notes_h = 126;
        $summary_w = 196;
        $notes_w = $content_width - $summary_w - 18;
        $summary_x = $margin + $notes_w + 18;
        $notes_y = $bottom_top - $notes_h;

        $stream .= sprintf("1 1 1 rg %.2F %.2F %.2F %.2F re f\n", $margin, $notes_y, $notes_w, $notes_h);
        $stream .= sprintf("%.3F %.3F %.3F RG %.2F %.2F %.2F %.2F re S\n", $border[0], $border[1], $border[2], $margin, $notes_y, $notes_w, $notes_h);
        $stream .= $this->text_cmd($margin + 14, $notes_y + $notes_h - 18, 'NOTAS', 'F2', 9, $primary);
        $note_text = trim((string) ($invoice['notes'] ?: 'Gracias por confiar en VexelDot.'));
        $ny = $notes_y + $notes_h - 36;
        foreach (array_slice($this->wrap_text($note_text, $notes_w - 28, 9), 0, 8) as $line) {
            $stream .= $this->text_cmd($margin + 14, $ny, $line, 'F1', 9, $text);
            $ny -= 11;
        }

        $stream .= sprintf("1 1 1 rg %.2F %.2F %.2F %.2F re f\n", $summary_x, $notes_y, $summary_w, $notes_h);
        $stream .= sprintf("%.3F %.3F %.3F RG %.2F %.2F %.2F %.2F re S\n", $border[0], $border[1], $border[2], $summary_x, $notes_y, $summary_w, $notes_h);
        $stream .= $this->text_cmd($summary_x + 14, $notes_y + $notes_h - 18, 'RESUMEN', 'F2', 9, $primary);

        $summary_rows = [
            ['Base imponible', $invoice['taxable_base']],
        ];
        if (($invoice['invoice_type'] ?? '') !== 'quote') {
            $summary_rows[] = ['IVA ' . number_format((float) ($invoice['tax_rate'] ?? 0), 0, ',', '.') . '%', $invoice['tax_amount']];
            if ((float) ($invoice['irpf_rate'] ?? 0) > 0) {
                $summary_rows[] = ['IRPF ' . number_format((float) ($invoice['irpf_rate'] ?? 0), 0, ',', '.') . '%', -1 * abs((float) ($invoice['irpf_amount'] ?? 0))];
            }
        }
        $ry = $notes_y + $notes_h - 36;
        foreach ($summary_rows as $row) {
            $stream .= $this->text_cmd($summary_x + 14, $ry, $row[0], 'F1', 9, $text);
            $stream .= $this->text_right_cmd($summary_x + $summary_w - 14, $ry, $this->fmt_money($row[1]), 'F2', 9, $text);
            $stream .= sprintf("%.3F %.3F %.3F RG %.2F %.2F m %.2F %.2F l S\n", $border[0], $border[1], $border[2], $summary_x + 14, $ry - 4, $summary_x + $summary_w - 14, $ry - 4);
            $ry -= 16;
        }

        $stream .= sprintf("%.3F %.3F %.3F rg %.2F %.2F %.2F 24 re f\n", $soft[0], $soft[1], $soft[2], $summary_x + 10, $notes_y + 10, $summary_w - 20);
        $stream .= $this->text_cmd($summary_x + 18, $notes_y + 18, 'TOTAL', 'F2', 12, $primary);
        $stream .= $this->text_right_cmd($summary_x + $summary_w - 18, $notes_y + 18, $this->fmt_money($invoice['total_amount']), 'F2', 11, $primary);

        $footer_y = 40;
        $stream .= sprintf("%.3F %.3F %.3F RG %.2F %.2F m %.2F %.2F l S\n", $border[0], $border[1], $border[2], $margin, $footer_y + 20, $width - $margin, $footer_y + 20);
        $pay_left = 'Forma de pago: Transferencia bancaria';
        if (!empty($settings['issuer_iban'])) {
            $pay_left .= '  |  IBAN: ' . $settings['issuer_iban'];
        }
        $right_footer = (($invoice['invoice_type'] ?? '') === 'quote') ? 'Documento comercial' : ('Estado VERI*FACTU: ' . ($invoice['verifactu_status'] ?: 'draft-ready'));
        $stream .= $this->text_cmd($margin, $footer_y + 6, $pay_left, 'F1', 8, $muted);
        $stream .= $this->text_cmd($margin, $footer_y - 6, 'Referencia: ' . $code, 'F1', 8, $muted);
        $stream .= $this->text_right_cmd($width - $margin, $footer_y + 6, $right_footer, 'F1', 8, $muted);

        $stream .= "Q\n";

        $font1 = $this->add_object("<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>");
        $font2 = $this->add_object("<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica-Bold >>");
        $content = $this->add_object("<< /Length " . strlen($stream) . " >>\nstream\n" . $stream . "\nendstream");

        $resources = "<< /Font << /F1 {$font1} 0 R /F2 {$font2} 0 R >>";
        if ($image_obj) {
            $resources .= " /XObject << /Im1 {$image_obj} 0 R >>";
        }
        $resources .= " >>";
        $page = $this->add_object("<< /Type /Page /Parent 0 0 R /MediaBox [0 0 {$width} {$height}] /Resources {$resources} /Contents {$content} 0 R >>");
        $pages = $this->add_object("<< /Type /Pages /Kids [{$page} 0 R] /Count 1 >>");
        $catalog = $this->add_object("<< /Type /Catalog /Pages {$pages} 0 R >>");
        $this->objects[$page - 1] = str_replace('/Parent 0 0 R', '/Parent ' . $pages . ' 0 R', $this->objects[$page - 1]);

        $pdf = "%PDF-1.4\n%âãÏÓ\n";
        $offsets = [0];
        foreach ($this->objects as $i => $obj) {
            $offsets[$i + 1] = strlen($pdf);
            $pdf .= ($i + 1) . " 0 obj\n" . $obj . "\nendobj\n";
        }
        $xref = strlen($pdf);
        $pdf .= "xref\n0 " . (count($this->objects) + 1) . "\n";
        $pdf .= "0000000000 65535 f \n";
        for ($i = 1; $i <= count($this->objects); $i++) {
            $pdf .= sprintf("%010d 00000 n \n", $offsets[$i]);
        }
        $pdf .= "trailer << /Size " . (count($this->objects) + 1) . " /Root {$catalog} 0 R >>\nstartxref\n{$xref}\n%%EOF";
        return $pdf;
    }
}
