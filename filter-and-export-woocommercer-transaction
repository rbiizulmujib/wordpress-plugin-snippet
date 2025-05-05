// Tambahkan menu baru di sidebar admin
add_action('admin_menu', 'menu_tabel_transaksi_wc');
function menu_tabel_transaksi_wc() {
    add_menu_page(
        'Tabel Transaksi',
        'Tabel Transaksi',
        'manage_woocommerce',
        'tabel-transaksi-wc',
        'halaman_tabel_transaksi_wc',
        'dashicons-cart',
        56
    );
}

// Fungsi untuk menampilkan halaman tabel transaksi
function halaman_tabel_transaksi_wc() {
    ?>
    <div class="wrap">
        <h1>Tabel Transaksi WooCommerce</h1>
        <?php
        // Ambil semua pesanan dengan status 'completed' atau 'processing'
        $args = array(
            'status' => array('completed', 'processing'),
            'limit' => -1,
            'return' => 'ids',
        );
        $orders = wc_get_orders($args);

        // Ambil daftar produk unik dari pesanan
        $produk_unik = array();
        foreach ($orders as $order_id) {
            $order = wc_get_order($order_id);
            foreach ($order->get_items() as $item) {
                $produk = $item->get_name();
                $produk_unik[$produk] = $produk;
            }
        }

        // Ambil produk yang dipilih dari parameter GET
        $selected_product = isset($_GET['produk']) ? sanitize_text_field($_GET['produk']) : '';

        // Form filter produk
        echo '<form method="GET" style="margin-bottom: 20px;">';
        echo '<input type="hidden" name="page" value="tabel-transaksi-wc">';
        echo '<select name="produk" onchange="this.form.submit()">';
        echo '<option value="">Semua Produk</option>';
        foreach ($produk_unik as $produk) {
            $selected = ($selected_product == $produk) ? 'selected' : '';
            echo "<option value='$produk' $selected>$produk</option>";
        }
        echo '</select>';
        echo '</form>';

        // Tabel transaksi
        echo '<table id="tabel-transaksi" class="widefat fixed striped">';
        echo '<thead><tr><th>Nama Produk</th><th>Nama Pembeli</th><th>Email Pembeli</th><th>Jumlah</th><th>Total</th><th>Tanggal</th></tr></thead>';
        echo '<tbody>';

        foreach ($orders as $order_id) {
            $order = wc_get_order($order_id);
            $order_date = $order->get_date_created()->date('Y-m-d');
            $customer_name = $order->get_billing_first_name() . ' ' . $order->get_billing_last_name();
            $email = $order->get_billing_email();

            foreach ($order->get_items() as $item) {
                $product_name = $item->get_name();
                if ($selected_product && $selected_product !== $product_name) continue;

                $quantity = $item->get_quantity();
                $total = wc_price($item->get_total());
                echo "<tr>
                    <td>{$product_name}</td>
                    <td>{$customer_name}</td>
                    <td>{$email}</td>
                    <td>{$quantity}</td>
                    <td>{$total}</td>
                    <td>{$order_date}</td>
                </tr>";
            }
        }

        echo '</tbody></table>';
        ?>
        <br>
        <button class="button" onclick="exportTableToCSV()">Export CSV</button>
        <button class="button button-primary" onclick="exportTableToExcel()">Export XLSX</button>
    </div>
    <script>
    function exportTableToCSV(filename = 'transaksi.csv') {
        let csv = [];
        const rows = document.querySelectorAll("table#tabel-transaksi tr");

        for (const row of rows) {
            const cols = row.querySelectorAll("td, th");
            const rowData = Array.from(cols).map(col => `"${col.innerText}"`);
            csv.push(rowData.join(","));
        }

        const blob = new Blob([csv.join("\n")], { type: 'text/csv' });
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement("a");
        a.setAttribute("href", url);
        a.setAttribute("download", filename);
        a.click();
    }

    function exportTableToExcel(filename = 'transaksi.xlsx') {
        const table = document.querySelector("#tabel-transaksi");
        const html = table.outerHTML.replace(/ /g, '%20');
        const data_type = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
        const a = document.createElement('a');
        a.href = 'data:' + data_type + ', ' + html;
        a.download = filename;
        a.click();
    }
    </script>
    <?php
}
