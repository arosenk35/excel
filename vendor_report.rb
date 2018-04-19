# frozen_string_literal: true

require 'rubygems'
require 'write_xlsx'
require 'pg'
include Writexlsx::Utility

class PrintInvDetails
  # connection = PG.connect(ENV.fetch('DW_DATABASE_URL'))

  dw_database_url = 'postgres://cangaroo:M7nEj4dN9UbcZkxE@trade-dwh.ckj3pghgzqvr.us-east-1.rds.amazonaws.com:5432/dwh'
  @connection = PG.connect(dw_database_url)
  @bill_status='Paid In Full'


  def self.generate_vendor_report(vendor,check_number,check_date,s_vendor_id,ns_vendor_id)
    # Create a new Excel workbook
    doc_vendor=vendor.upcase
    doc_vendor.gsub!(/[^0-9A-Za-z]/, '')
    doc_date= if check_date.nil? then Date.today.to_s else check_date[0..9] end
    doc_type= if @bill_status=='Open' then 'PRELIM' else 'TRADE' end
    docname = doc_type + '_'+ s_vendor_id + '_' + doc_vendor + '_' + doc_date + '.xlsx'
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

    print_remittance_summary(workbook, ns_vendor_id,check_number)
    print_remittance(workbook, ns_vendor_id,check_number)

    workbook.close
  end

  # Remittance Datilas Report
  def self.print_remittance(workbook, vendor, check_number)
    data = get_remittance_details(@connection, vendor, check_number)
    return unless !data.nil?
    worksheet = workbook.add_worksheet('Remittance Details')

    layout = [
      { field: 'charge_type', offset: 0, header: 'Transaction Type' },
      { field: 'invoice_id', offset: 0, header: 'Trade Invoice Id' },
      { field: 'invoice_date', offset: 0, header: 'Shipment Date', format: @date_format },
      { field: 'week', offset: 0, header: 'Week Nbr' },
      { field: 'firstname', offset: 0, header: 'Fist Name' },
      { field: 'lastname', offset: 0, header: 'Last Name' },
      { field: 'shipment_number', offset: 0, header: 'Roaster Order Nbr' },
      { field: 'description', offset: 0, header: 'Description' },
      { field: 'tracking_url', offset: 0, header: 'Tracking', format: @url_format },
      { field: 'bill_qty', offset: 0, header: 'Bags/Qty', total: true },
      { field: 'unit_cost', offset: 0, header: 'Unit Cost', format: @money_format, type: 'float' },
      { field: 'bill_amt', offset: 0, header: 'Gross', format: @money_format, type: 'float', total: true }
    ]

    vert_layout = [
      { field: 'vendor', offset: 0, header: 'Roaster:' },
      { field: 'check_date', offset: 0, header: 'Payment Date:' }
    ]

    @row=0
    worksheet.insert_image('A1','..\cangaroo\app\excel\trade-dark.jpg')
    @row+=3

    print_vertical_header(data, vert_layout, worksheet)
    print_std_report(data, layout, worksheet)
  end

  #remmitance summary report
  def self.print_remittance_summary(workbook, vendor,check_number)
    data = get_remittance_summary(@connection, vendor,check_number)
    return unless !data.nil?
    worksheet = workbook.add_worksheet('Remittance Summary')
    layout = [
      { field: 'charge_type', offset: 0, header: 'Transaction Type' },
      { field: 'description', offset: 0, header: 'Description' },
      { field: 'bill_qty', offset: 0, header: 'Bags/Qty', total: true },
      { field: 'unit_cost', offset: 0, header: 'Unit Cost', format: @money_format, type: 'float' },
      { field: 'bill_amt', offset: 0, header: 'Gross', format: @money_format, type: 'float', total: true }
    ]

    vert_layout = [
      { field: 'vendor', offset: 0, header: 'Roaster:' },
      { field: 'check_date', offset: 0, header: 'Payment Date:' }
    ]

    @row=0
    worksheet.insert_image('A1','..\cangaroo\app\excel\trade-dark.jpg')
    @row+=3
    print_vertical_header(data, vert_layout, worksheet)
    print_std_report(data, layout, worksheet)
  end

  # standard report layout emgine
  def self.print_std_report(data, layout, worksheet)
    col =0
    # print header
    print_header(layout, worksheet)
    # print lines
    data.each_with_index do |record, index|
      print_line(record, @row, layout, worksheet)
      @row+=1
      end

    # print totals
    print_totals(layout, worksheet)
  end

  def self.print_header(layout, worksheet)
    offset=0
    layout.each_with_index do |column, index|
      offset += column[:offset].to_i
      worksheet.write(@row, index + offset, column[:header], @format_heading)
    end
    @row +=1
  end

  def self.print_vertical_header(data, layout, worksheet)
    if @bill_status=='Open'
      worksheet.write(@row, 4, 'Prelim Remittance',  @format_total)
    end

    layout.each do |column|
      worksheet.write(@row, 0, column[:header], @format_heading)
      worksheet.write(@row, 1, (column[:type] == 'float' ? data.first[column[:field]].to_f : data.first[column[:field]]), column[:format])
     @row+=1
    end
    @row+=1
  end

  def self.print_line(record, row, layout, worksheet)
    offset=0
    layout.each_with_index do |column, index|
      offset += column[:offset].to_i
      worksheet.write(row, index + offset, (column[:type] == 'float' ? record[column[:field]].to_f : record[column[:field]]), column[:format])
    end
  end

  def self.print_totals(layout, worksheet)
    start = @row + 1
    offset=0
    layout.each_with_index do |column, index|
      offset += column[:offset].to_i
      next unless column[:total] == true
      worksheet.write(start, 0, 'Totals', @format_total)
      sum = '=SUM(' + xl_rowcol_to_cell(1, index + offset) + ':' + xl_rowcol_to_cell(@row, index + offset) + ')'
      worksheet.write(start, index + offset, sum, column[:format])
    end
  end

  def self.get_remittance_details(connection, vendor,check_number)
    sql = <<-eosql
                SELECT  vendor, charge_type, check_number, to_char(check_date,'mm/dd/yyyy') check_date, invoice_id,
                        to_char(invoice_date,'mm/dd/yyyy') invoice_date,
                        date_part('week',invoice_date) week,
                        check_amt, sku, description, bill_qty, bill_amt, shipment_qty,
                        unit_cost::numeric, order_number, shipment_number, firstname,
                        lastname, tracking_url
                    FROM public.netsuite_remittance_details_vw
                        where ns_vendor_id='#{vendor}' and
                        (check_number='#{check_number}' or ( check_number is null and '#{check_number}'='')) and
                        bill_status='#{@bill_status}'
                        order by 1,2,3,4,5,6
            eosql
    connection.exec sql
  end

  def self.get_remittance_summary(connection, vendor,check_number)
    sql = <<-eosql
                SELECT  vendor, charge_type,check_number, to_char(check_date,'mm/dd/yyyy') check_date,
               description,unit_cost::numeric,
                    sum(bill_qty) bill_qty, sum(bill_amt) bill_amt
                    FROM public.netsuite_remittance_details_vw
                        where ns_vendor_id='#{vendor}' and
                        (check_number='#{check_number}' or (check_number is null and '#{check_number}'='')) and
                        bill_status='#{@bill_status}'
                        group by 1,2,3,4,5,6
                        order by 1,2,3,4,5,6
            eosql
    connection.exec sql
  end

  def self.get_remittance_vendor(connection)
    sql = <<-eosql
    SET CLIENT_ENCODING TO 'utf8';
                SELECT  distinct r.vendor, r.check_number,r.check_date, r.s_vendor_id,r.ns_vendor_id
                    FROM public.netsuite_remittance_details_vw r
                    where bill_status='#{@bill_status}'
                    -----and check_number not in (select check_number from table.emailed e where r.check_number=e.check_number)
            eosql
    connection.exec sql
  end

  def self.generate_remittance(type)
    @bill_status =if type.downcase == 'remittance'
                    'Paid In Full'
                  else
                    'Open'
                  end

      data=get_remittance_vendor(@connection)
      data.each do |record|
        ### need to only print dropship vendors .... requires fix
        #escape quote for postgress
        vendor=record['vendor']
        check_number=record['check_number']
        check_date=record['check_date']
        s_vendor_id=record['s_vendor_id']
        ns_vendor_id=record['ns_vendor_id']
        generate_vendor_report(vendor, check_number,check_date,s_vendor_id,ns_vendor_id)
      end
  end
end
###chcp is windows machine issue
`chcp 65001`
#PrintInvDetails.generate_remittance('prelim')
PrintInvDetails.generate_remittance('remittance')