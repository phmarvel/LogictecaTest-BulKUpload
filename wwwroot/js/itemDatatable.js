$(document).ready(function () {

    $('#itemDatatable tfoot th').each(function () {
        var title = $(this).text();
        $(this).html('<input type="text" placeholder="Search ' + title + '" />');
    });

    var table =   $("#itemDatatable").DataTable({
           "processing": true,
           "filter": true,
           serverSide:true,
       "ajax": {
           "dataType": 'json',
           "type": "POST",
           "url": "/api/item",
           "contentType": 'application/x-www-form-urlencoded',
        },
        "columnDefs": [{
            "targets": [0],
            "visible": false,
            "searchable": false
        }],
        initComplete: function () {
            // Apply the search
            this.api().columns().every(function () {
                var that = this;

                $('input', this.footer()).on('keyup change clear', function () {
                    if (that.search() !== this.value) {
                        that
                            .search(this.value)
                            .draw();
                    }
                });
            });
        },
        "columns": [
            {"data": "band", "name": "Band", "autoWidth": true },
            {"data": "category_Code", "name": "Category Code", "autoWidth": true },
            {"data": "manufacturer", "name": "Manufacturer", "autoWidth": true },
            { "data": "part_SKU", "name": "Part SKU", "autoWidth": true },
            { "data": "item_Description", "name": "Item Description", "autoWidth": true },
            { "data": "list_Price", "name": "List Price", "autoWidth": true },
            { "data": "minimum_Discount", "name": "Minimum Discount", "autoWidth": true },
            { "data": "discounted_Price", "name": "Discounted Price", "autoWidth": true },

        ],
        dom: 'Bfrtip',
        "buttons": [
            {
                text: 'Export',
                className: "btn btn-primary",
                action: function (e, dt, node, config) {
                    console.log(dt.ajax.params())
                    debugger
                    $.ajax({
                        "url": "/api/item/Export",
                        "dataType": 'json',
                        "type": "POST",
                        "contentType": 'application/x-www-form-urlencoded',
                        "data": dt.ajax.params(),
                        "success": function (res, status, xhr) {
                            debugger
                            window.open(window.location.origin  +res.url, '_blank');

                        }
                    });
                }
            }]
    });


    $("#ImportBtn").change(function () {
        if (this.files.length) {
            var file = this.files[0];
            var formData = new FormData();
            formData.append('File', file);
            $.ajax({
                url: '/api/item/Import',  //Server script to process data
                type: 'POST',
                data: formData,
                contentType: false,
                processData: false,
                //Ajax events
                success: function (html) {
                    debugger
                    alert("Import Process Start Successfully");
                    window.location.reload()
                }
            });
        }
    })
});