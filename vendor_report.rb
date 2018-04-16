# frozen_string_literal: true

require 'rubygems'
require 'write_xlsx'
require 'pg'
require 'date'
require 'byebug'
include Writexlsx::Utility

class PrintInvDetails
  # connection = PG.connect(ENV.fetch('DW_DATABASE_URL'))

  dw_database_url = 'postgres://cangaroo:M7nEj4dN9UbcZkxE@trade-dwh.ckj3pghgzqvr.us-east-1.rds.amazonaws.com:5432/dwh'
  @connection = PG.connect(dw_database_url)

  def self.generate_vendor_report(vendor)
    # Create a new Excel workbook
    docname = vendor + Date.today.to_s + '.xlsx'
    workbook = WriteXLSX.new(docname)

    @format_total = workbook.add_format(
      bold: 1,
      border: 1
    )

    @format_heading = workbook.add_format(
      bold: 1,
      fg_color: 'silver',
      pattern: 1,
      border: 1
    )
    @url_format = workbook.add_format(
      color: 'blue',
      underline: 1
    )
    @money_format = workbook.add_format(num_format: '$#,##0.00')
    @date_format  = workbook.add_format(num_format: 'mm/dd/yyyy', align: 'left')

    print_remittance(workbook, vendor)
    print_remittance_summary(workbook, vendor)

    workbook.close
  end

  # Remittance Datilas Report
  def self.print_remittance(workbook, vendor)
    data = get_remittance_details(@connection, vendor)
    worksheet = workbook.add_worksheet('Remittance Details')

    layout = [
      { field: 'vendor', offset: 0, header: 'Supplier' },
      { field: 'check_date', offset: 0, header: 'Check Date' },
      { field: 'check_number', offset: 0, header: 'ACH/Chk Ref' },
      { field: 'charge_type', offset: 0, header: 'Charge Type' },
      { field: 'invoice_id', offset: 0, header: 'Invoice Id' },
      { field: 'invoice_date', offset: 0, header: 'Shipment Date', format: @date_format },
      { field: 'week', offset: 0, header: 'Week Nbr' },
      { field: 'firstname', offset: 0, header: 'Fist Name' },
      { field: 'lastname', offset: 0, header: 'Last Name' },
      { field: 'order_number', offset: 0, header: 'Order Nbr' },
      { field: 'shipment_number', offset: 0, header: 'Shipment Nbr' },
      { field: 'sku', offset: 2, header: 'sku' },
      { field: 'unit_cost', offset: 0, header: 'Unit Cost', format: @money_format, type: 'float' },
      { field: 'description', offset: 0, header: 'Description' },
      { field: 'tracking_url', offset: 0, header: 'Tracking', format: @url_format },
      { field: 'bill_amt', offset: 0, header: 'Gross', format: @money_format, type: 'float', total: true },
      { field: 'bill_qty', offset: 0, header: 'Bags/Qty', total: true }
    ]
    print_std_report(data, layout, worksheet)
  end

  #remmitance summary report
  def self.print_remittance_summary(workbook, vendor)
    data = get_remittance_summary(@connection, vendor)
    worksheet = workbook.add_worksheet('Remmitance Summary')
    layout = [
      { field: 'vendor', offset: 0, header: 'Supplier' },
      { field: 'charge_type', offset: 0, header: 'Charge Type' },
      { field: 'week', offset: 0, header: 'Week Nbr' },
      { field: 'bill_amt', offset: 0, header: 'Gross', format: @money_format, type: 'float', total: true },
      { field: 'bill_qty', offset: 0, header: 'Bags/Qty', total: true }
    ]

    print_std_report(data, layout, worksheet)
  end

  # standard report layout emgine
  def self.print_std_report(data, layout, worksheet)
    col = row = 0
    # print header
    print_header(0, layout, worksheet)
    last_row = 1

    # print lines
    start = last_row
    data.each_with_index do |record, index|
      print_line(record, start + index, layout, worksheet)
      last_row = start + index
    end

    # print totals
    print_totals(last_row, layout, worksheet)
  end

  def self.print_header(row, layout, worksheet)
    offset=0
    layout.each_with_index do |column, index|
      offset += column[:offset].to_i
      worksheet.write(row, index + offset, column[:header], @format_heading)
    end
  end

  def self.print_line(record, row, layout, worksheet)
    offset=0
    layout.each_with_index do |column, index|
      offset += column[:offset].to_i
      worksheet.write(row, index + offset, (column[:type] == 'float' ? record[column[:field]].to_f : record[column[:field]]), column[:format])
    end
  end

  def self.print_totals(last_row, layout, worksheet)
    start = last_row + 2
    offset=0
    layout.each_with_index do |column, index|
      offset += column[:offset].to_i
      next unless column[:total] == true
      worksheet.write(start, 0, 'Totals', @format_total)
      sum = '=SUM(' + xl_rowcol_to_cell(1, index + offset) + ':' + xl_rowcol_to_cell(last_row, index + offset) + ')'
      worksheet.write(start, index + offset, sum, column[:format])
    end
  end

  def self.get_remittance_details(connection, vendor)
    sql = <<-eosql
                SELECT  vendor, charge_type, check_number, check_date, invoice_id,
                        to_char(invoice_date,'mm/dd/yyyy') invoice_date,
                        date_part('week',invoice_date) week,
                        check_amt, sku, description, bill_qty, bill_amt, shipment_qty,
                        unit_cost::numeric, order_number, shipment_number, firstname,
                        lastname, tracking_url
                    FROM public.netsuite_remittance_details_vw
                        where vendor='#{vendor}'
                        order by 1,2,3,4,5,6
            eosql
    connection.exec sql
  end

  def self.get_remittance_summary(connection, vendor)
    sql = <<-eosql
                SELECT  vendor, charge_type,
                    date_part('week',invoice_date) week,
                    sum(bill_qty) bill_qty, sum(bill_amt) bill_amt
                    FROM public.netsuite_remittance_details_vw
                        where vendor='#{vendor}'
                        group by 1,2,3
                        order by 1,2,3
            eosql
    connection.exec sql
  end

  def self.get_remittance_vendor(connection)
    sql = <<-eosql
                SELECT  distinct vendor, check_number, check_date
                    FROM public.netsuite_remittance_details_vw
                        --where not printed
            eosql
    connection.exec sql
  end

  def self.generate_remittance
      data=get_remittance_vendor(@connection)
      data.each do |record|
        ### need to only print drpship vendors .... requires fix
        #escpape quote for postgress
        vendor=record['vendor'].gsub("'","''")
        generate_vendor_report(vendor)
      end
  end
end

PrintInvDetails.generate_remittance