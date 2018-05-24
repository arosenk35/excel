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

  #
  # Main ...gerenate the various reports/woksheets for each payment/vendor
  #
  def self.generate_vendor_report(vendor,
                                  ns_vendor_id)
    # Create a new Excel workbook
    doc_vendor=vendor.upcase
    doc_vendor.gsub!(/[^0-9A-Za-z]/, '')
    doc_date= if @check_date.nil? then Date.today.to_s else @check_date[0..9] end
    doc_type= if @type=='remittance' then 'TRADE' else 'PRELIM'  end
    docname = doc_type + '_'+ @s_vendor_id + '_' + doc_vendor + '_' + doc_date + '.xlsx'
    workbook = WriteXLSX.new(docname)

    #set std formats
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

    #print reports
    print_remittance_summary(workbook, ns_vendor_id)
    print_remittance(workbook, ns_vendor_id)

    workbook.close

    #mark paymentts emailed
    if @type.downcase == 'remittance'
      mark_remittance_emailed(@connection,ns_vendor_id)
    end

  end
  #
  # Remittance Details Report
  #
  def self.print_remittance(workbook, vendor)
    data = get_remittance_details(@connection, vendor)
    return unless !data.nil?
    worksheet = workbook.add_worksheet('Remittance Details')

    layout = [
      { field: 'charge_type', offset: 0, header: 'Transaction Type' },
      { field: 'invoice_id', offset: 0, header: 'Trade Invoice Id' },
      { field: 'invoice_date', offset: 0, header: 'Shipment Date', format: @date_format },
      { field: 'week', offset: 0, header: 'Week Nbr' },
      { field: 'firstname', offset: 0, header: 'Fist Name' },
      { field: 'lastname', offset: 0, header: 'Last Name' },
      { field: 'batch_id', offset: 0, header: 'Roaster Batch Id' },
      { field: 'shipment_number', offset: 0, header: 'Roaster Order Nbr' },
      { field: 'description', offset: 0, header: 'Description' },
      { field: 'tracking_url', offset: 0, header: 'Tracking', format: @url_format },
      { field: 'bill_qty', offset: 0, header: 'Bags/Qty', total: true },
      { field: 'unit_cost', offset: 0, header: 'Unit Cost', format: @money_format, type: 'float' },
      { field: 'bill_amt', offset: 0, header: 'Gross', format: @money_format, type: 'float', total: true }
    ]

    vert_layout = [
      { field: 'vendor', offset: 0, header: 'Roaster:' },
      { field: '@max_inv_date', offset: 0, header: 'Invoices Paid Thru:' , type: 'var' },
      { field: '@max_check_date', offset: 0, header: 'Payment Date:' ,type:'var'}
    ]

    @row=0
    worksheet.insert_image('A1','..\excel\trade-dark.jpg')
    @row+=3

    print_vertical_header(data, vert_layout, worksheet)
    print_std_report(data, layout, worksheet)
  end
  #
  #remmitance summary report
  #
  def self.print_remittance_summary(workbook, vendor)
    data = get_remittance_summary(@connection, vendor)
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
      { field: '@max_inv_date', offset: 0, header: 'Invoices Paid Thru:',type:'var' },
      { field: '@max_check_date', offset: 0, header: 'Payment Date:' ,type:'var'}
    ]

    @row=0
    worksheet.insert_image('A1','..\excel\trade-dark.jpg')
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
      worksheet.write(@row, 1,
                        (if column[:type] == 'float'
                            data.first[column[:field]].to_f
                          elsif column[:type] == 'var'
                            eval(column[:field])
                          else
                            data.first[column[:field]]
                        end),
                        column[:format])
     @row+=1
    end
    @row+=1
  end

  def self.print_line(record, row, layout, worksheet)
    offset=0
    layout.each_with_index do |column, index|
      offset += column[:offset].to_i
      worksheet.write(row, index + offset, (if column[:type] == 'float'
                                                record[column[:field]].to_f
                                            elsif column[:type] == 'var'
                                                 eval(column[:field])
                                            else
                                              record[column[:field]]
                                            end), column[:format])
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

  def self.get_remittance_details(connection, vendor)
    sql = <<-eosql
                SELECT  vendor, charge_type, invoice_id,
                        to_char(invoice_date,'mm/dd/yyyy') invoice_date,
                        date_part('week',invoice_date) week,
                        check_amt, sku, description, bill_qty, bill_amt, shipment_qty,
                        unit_cost::numeric, order_number, shipment_number, firstname,
                        lastname, tracking_url,batch_id
                    FROM public.netsuite_remittance_details_vw r
                        left join cangaroo_interface.ap_emailed_remittances e on  r.payment_id=e.payment_id
                        where ns_vendor_id='#{vendor}'
                        and e.payment_id is null
                        and bill_status in (#{@bill_status})
                        order by 1,2,3,4,5
            eosql
    connection.exec sql
  end

  def self.get_remittance_summary(connection, vendor)
    sql = <<-eosql
                SELECT  vendor, charge_type,
               description,unit_cost::numeric,
                    sum(bill_qty) bill_qty, sum(bill_amt) bill_amt
                    FROM public.netsuite_remittance_details_vw r
                        left join cangaroo_interface.ap_emailed_remittances e on  r.payment_id=e.payment_id
                        where ns_vendor_id='#{vendor}'
                        and e.payment_id is null
                        and bill_status in (#{@bill_status})
                        group by 1,2,3,4
                        order by 1,2,3,4
            eosql
    connection.exec sql
  end

  def self.get_remittance_vendor(connection)
    #only get payments that have not been emailed
    sql = <<-eosql
    SET CLIENT_ENCODING TO 'utf8';
                SELECT
                r.vendor,
                r.s_vendor_id,
                r.ns_vendor_id,
                to_char(max(invoice_date),'mm/dd/yyyy') max_inv_date,
                to_char(max(check_date),'mm/dd/yyyy') max_check_date
                    FROM public.netsuite_remittance_details_vw r
                         left join cangaroo_interface.ap_emailed_remittances e on  r.payment_id=e.payment_id
                         where bill_status in (#{@bill_status})
                         and e.payment_id is null
                group by r.vendor, r.s_vendor_id,r.ns_vendor_id
            eosql
    connection.exec sql
  end

  #trap invoices that have been emailed .... we do not want to create and resend them
  def self.mark_remittance_emailed(connection,payment_id)
    sql = <<-eosql
            insert into cangaroo_interface.ap_emailed_remittances (payment_id) values('#{payment_id}')
            eosql
    return unless !payment_id.nil?
    connection.exec sql
  end

  #
  #loop thru open/nom emailed invoices
  #
  def self.generate_remittance(type)
    #set runtime options
    # "'Open','Paid In Full'"
    @bill_status =if type.downcase == 'remittance'
                    "'Paid In Full'"
                  else
                    "'Open'"
                  end
    @type=type

      data=get_remittance_vendor(@connection)
      data.each do |record|
        # we only only print where vendor_category = 7/inventory(=dropship) all other statements run the other way (locked in sql view)
        # use payment id not check as wire transfer/on deposits do not fill in the check number
        vendor=record['vendor']
        payment_id=record['payment_id']
        ns_vendor_id=record['ns_vendor_id']

        #document level variables
        @check_date=record['check_date']
        @s_vendor_id=record['s_vendor_id']
        @max_inv_date=record['max_inv_date']
        @max_check_date=record['max_check_date']

        generate_vendor_report(vendor,
                              ns_vendor_id)
      end
  end
end
###chcp is windows machine issue
`chcp 65001` if Gem.win_platform?
#PrintInvDetails.generate_remittance('prelim')
PrintInvDetails.generate_remittance('remittance')