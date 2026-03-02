[
    {
        "table_name": "clients",
        "column_name": "id",
        "data_type": "uuid",
        "is_nullable": "NO",
        "column_default": "gen_random_uuid()"
    },
    {
        "table_name": "clients",
        "column_name": "nom",
        "data_type": "text",
        "is_nullable": "YES",
        "column_default": null
    },
    {
        "table_name": "clients",
        "column_name": "created_at",
        "data_type": "timestamp with time zone",
        "is_nullable": "YES",
        "column_default": "now()"
    },
    {
        "table_name": "clients",
        "column_name": "siren",
        "data_type": "text",
        "is_nullable": "YES",
        "column_default": null
    },
    {
        "table_name": "clients",
        "column_name": "pennylane_id",
        "data_type": "bigint",
        "is_nullable": "NO",
        "column_default": null
    },
    {
        "table_name": "customer_pricing",
        "column_name": "id",
        "data_type": "uuid",
        "is_nullable": "NO",
        "column_default": "uuid_generate_v4()"
    },
    {
        "table_name": "customer_pricing",
        "column_name": "client_id",
        "data_type": "uuid",
        "is_nullable": "YES",
        "column_default": null
    },
    {
        "table_name": "customer_pricing",
        "column_name": "product_id",
        "data_type": "uuid",
        "is_nullable": "YES",
        "column_default": null
    },
    {
        "table_name": "customer_pricing",
        "column_name": "custom_price_ht",
        "data_type": "numeric",
        "is_nullable": "NO",
        "column_default": null
    },
    {
        "table_name": "customer_pricing",
        "column_name": "currency",
        "data_type": "text",
        "is_nullable": "YES",
        "column_default": "'EUR'::text"
    },
    {
        "table_name": "customer_pricing",
        "column_name": "updated_at",
        "data_type": "timestamp with time zone",
        "is_nullable": "YES",
        "column_default": "now()"
    },
    {
        "table_name": "invoice_logs",
        "column_name": "id",
        "data_type": "uuid",
        "is_nullable": "NO",
        "column_default": "uuid_generate_v4()"
    },
    {
        "table_name": "invoice_logs",
        "column_name": "client_id",
        "data_type": "uuid",
        "is_nullable": "YES",
        "column_default": null
    },
    {
        "table_name": "invoice_logs",
        "column_name": "total_ht",
        "data_type": "numeric",
        "is_nullable": "YES",
        "column_default": null
    },
    {
        "table_name": "invoice_logs",
        "column_name": "status",
        "data_type": "text",
        "is_nullable": "YES",
        "column_default": "'draft'::text"
    },
    {
        "table_name": "invoice_logs",
        "column_name": "sent_at",
        "data_type": "timestamp with time zone",
        "is_nullable": "YES",
        "column_default": null
    },
    {
        "table_name": "invoice_logs",
        "column_name": "pennylane_ref",
        "data_type": "text",
        "is_nullable": "YES",
        "column_default": null
    },
    {
        "table_name": "products",
        "column_name": "id",
        "data_type": "uuid",
        "is_nullable": "NO",
        "column_default": "uuid_generate_v4()"
    },
    {
        "table_name": "products",
        "column_name": "label",
        "data_type": "text",
        "is_nullable": "NO",
        "column_default": null
    },
    {
        "table_name": "products",
        "column_name": "price",
        "data_type": "numeric",
        "is_nullable": "YES",
        "column_default": null
    },
    {
        "table_name": "products",
        "column_name": "vat_rate",
        "data_type": "text",
        "is_nullable": "YES",
        "column_default": null
    },
    {
        "table_name": "products",
        "column_name": "pennylane_id",
        "data_type": "bigint",
        "is_nullable": "NO",
        "column_default": null
    },
    {
        "table_name": "products",
        "column_name": "created_at",
        "data_type": "timestamp with time zone",
        "is_nullable": "YES",
        "column_default": "now()"
    },
    {
        "table_name": "products",
        "column_name": "price_before_tax",
        "data_type": "numeric",
        "is_nullable": "YES",
        "column_default": null
    },
    {
        "table_name": "products",
        "column_name": "unit",
        "data_type": "text",
        "is_nullable": "YES",
        "column_default": null
    },
    {
        "table_name": "products",
        "column_name": "description",
        "data_type": "text",
        "is_nullable": "YES",
        "column_default": null
    },
    {
        "table_name": "products",
        "column_name": "archived_at",
        "data_type": "timestamp with time zone",
        "is_nullable": "YES",
        "column_default": null
    },
    {
        "table_name": "products",
        "column_name": "updated_at",
        "data_type": "timestamp with time zone",
        "is_nullable": "YES",
        "column_default": null
    }
]