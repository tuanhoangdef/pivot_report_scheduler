import ast
import base64
from collections import deque
import io
from numbers import Number

from dateutil.relativedelta import relativedelta
from odoo import api, fields, models
from odoo.tools import osutil
import xlsxwriter
from odoo.tools.safe_eval import datetime, safe_eval


class ScheduledPivotReport(models.Model):
    _name = 'scheduled.pivot.report'
    _description = 'Scheduled Pivot Report'

    name = fields.Char(required=True)
    active = fields.Boolean(default=True)
    schedule = fields.Selection(
        [('daily', 'Daily'), ('weekly', 'Weekly')],
        required=True,
        default='daily',
    )
    favorite_filter_id = fields.Many2one(
        'ir.filters',
        string='Favorite Filter',
        required=True,
        domain="[('model_id', '!=', False)]",
    )
    model_id = fields.Char(
        string='Model',
        compute='_compute_model_id',
        store=True,
    )
    title = fields.Char(
        help='Used as the worksheet name, email subject, and attachment prefix. '
             'When empty, the favorite filter name is used.'
    )
    email_template_id = fields.Many2one(
        'mail.template',
        string='Email Template',
        domain="[('model', '=', 'scheduled.pivot.report')]",
        default=lambda self: self.env.ref(
            'pivot_report_scheduler.email_template_scheduled_pivot_report',
            raise_if_not_found=False,
        ),
    )
    recipient_partner_ids = fields.Many2many(
        'res.partner',
        'scheduled_pivot_report_partner_rel',
        'report_id',
        'partner_id',
        string='Recipients',
    )
    recipient_emails = fields.Char(
        help='Extra comma-separated email addresses.',
    )
    weekly_send_day = fields.Selection(
        [
            ('0', 'Monday'),
            ('1', 'Tuesday'),
            ('2', 'Wednesday'),
            ('3', 'Thursday'),
            ('4', 'Friday'),
            ('5', 'Saturday'),
            ('6', 'Sunday'),
        ],
        default='0',
        required=True,
    )

    @api.depends('favorite_filter_id.model_id')
    def _compute_model_id(self):
        for report in self:
            report.model_id = report.favorite_filter_id.model_id

    def _get_report_user(self):
        self.ensure_one()
        favorite_filter = self.favorite_filter_id
        if 'user_id' in favorite_filter._fields:
            return favorite_filter.user_id or favorite_filter.create_uid or self.env.user
        if favorite_filter.user_ids:
            return (
                favorite_filter.user_ids.filtered(lambda user: user == self.env.user)
                or favorite_filter.user_ids[:1]
            )
        return favorite_filter.create_uid or self.env.user

    def _get_favorite_domain(self):
        self.ensure_one()
        # ir.filters domains can contain Odoo helper expressions such as
        # context_today() and relativedelta(), so evaluate them with a narrow context.
        return safe_eval(self.favorite_filter_id.domain, self._get_eval_context())

    def _get_eval_context(self):
        return {
            'datetime': datetime,
            'context_today': datetime.datetime.now,
            'relativedelta': relativedelta,
        }

    def _get_favorite_context(self):
        self.ensure_one()
        return ast.literal_eval(self.favorite_filter_id.context)

    def _get_group_key(self, line, groupby):
        if not groupby:
            return 'total', 'Total'
        value = line.get(groupby)
        if isinstance(value, (tuple, list)):
            return value[0] or False, value[1] or 'None'
        if value in (False, None):
            return False, 'None'
        return value, str(value)

    def _get_axis_info(self, line, groupbys):
        # Convert a read_group result into the raw key used for lookups and the
        # display labels used to build row/column headers.
        keys = []
        labels = []
        for groupby in groupbys:
            key, label = self._get_group_key(line, groupby)
            keys.append(key)
            labels.append(label)
        return tuple(keys) or ('total',), tuple(labels) or ('Total',)

    def _format_measure_label(self, model, measure):
        if measure == '__count':
            return 'Count'
        field = model._fields.get(measure)
        return field.string if field else measure

    def _get_measure_fields(self, model, measures):
        # read_group expects aggregate specs such as "price_total:sum".
        # Keep the original measure names separately because read_group returns
        # values under those original names.
        measure_fields = []
        valid_measures = []
        for measure in measures:
            if measure == '__count':
                measure_fields.append(measure)
                valid_measures.append(measure)
                continue
            field = model._fields.get(measure)
            if not field or not field.aggregator:
                continue
            measure_fields.append('%s:%s' % (measure, field.aggregator))
            valid_measures.append(measure)
        if not valid_measures:
            return [], ['__count']
        return measure_fields, valid_measures

    def _read_pivot_matrix(self):
        self.ensure_one()
        favorite_filter_context = self._get_favorite_context()
        model = self.env[self.model_id].with_user(self._get_report_user()).with_context(**favorite_filter_context)

        # The favorite filter is the source of truth for pivot axes and measures,
        # matching how Odoo restores a saved pivot from ir.filters.context.
        row_groupbys = favorite_filter_context.get('pivot_row_groupby') or favorite_filter_context.get('group_by', [])
        column_groupbys = favorite_filter_context.get('pivot_column_groupby', [])
        measures = favorite_filter_context.get('pivot_measures', []) or ['__count']
        domain = self._get_favorite_domain()
        fields_list, measures = self._get_measure_fields(model, measures)
        rows = []
        seen_rows = set()
        columns = {}
        row_values = {}
        row_totals = {}
        column_totals = {}
        grand_totals = {}

        # Query each row depth separately so parent rows use their own database
        # aggregation. This is required for non-additive measures like averages.
        row_depths = range(1, len(row_groupbys) + 1) if row_groupbys else [0]
        for depth in row_depths:
            depth_row_groupbys = row_groupbys[:depth]
            lines = model.read_group(domain, fields_list, depth_row_groupbys + column_groupbys, lazy=False)
            for line in lines:
                row_keys, row_labels = self._get_axis_info(line, depth_row_groupbys)
                col_keys, col_labels = self._get_axis_info(line, column_groupbys)
                if depth == len(row_groupbys):
                    # Only the deepest row grouping defines visible leaf columns
                    # and lets us emit the full row tree in Odoo pivot order.
                    columns[col_keys] = col_labels
                    for level in range(depth):
                        tree_key = row_keys[:level + 1]
                        if tree_key not in seen_rows:
                            rows.append((tree_key, row_labels[level], level + 1))
                            seen_rows.add(tree_key)
                for measure in measures:
                    row_values[(row_keys, col_keys, measure)] = line.get(measure, 0) or 0
            # Row totals are queried without column groupbys so Odoo computes
            # the correct total/average for the whole row.
            for line in model.read_group(domain, fields_list, depth_row_groupbys, lazy=False):
                row_keys, _row_labels = self._get_axis_info(line, depth_row_groupbys)
                for measure in measures:
                    row_totals[(row_keys, measure)] = line.get(measure, 0) or 0

        # Column totals are the symmetric case: group only by column fields.
        for line in model.read_group(domain, fields_list, column_groupbys, lazy=False):
            col_keys, col_labels = self._get_axis_info(line, column_groupbys)
            for measure in measures:
                column_totals[(col_keys, measure)] = line.get(measure, 0) or 0

        # Empty groupby means one aggregate row over the full filtered dataset.
        for line in model.read_group(domain, fields_list, [], lazy=False):
            for measure in measures:
                grand_totals[measure] = line.get(measure, 0) or 0

        # Adjacent equal prefixes are required for merged column headers.
        columns = sorted(columns.items(), key=lambda item: item[0])
        return {
            'model': model,
            'measures': measures,
            'measure_labels': {
                measure: self._format_measure_label(model, measure)
                for measure in measures
            },
            'column_group_depth': len(column_groupbys),
            'rows': rows,
            'columns': columns,
            'row_values': row_values,
            'row_totals': row_totals,
            'column_totals': column_totals,
            'grand_totals': grand_totals,
        }

    def _get_pivot_export_data(self, matrix):
        # Build the same JSON-like structure that Odoo's pivot XLSX controller
        # receives from the frontend PivotModel.exportData().
        columns = matrix['columns']
        rows = matrix['rows']
        row_values = matrix['row_values']
        measures = matrix['measures']
        measure_labels = matrix['measure_labels']
        has_total_column = len(columns) > 1
        column_totals = matrix['column_totals']
        row_totals = matrix['row_totals']
        grand_totals = matrix['grand_totals']
        col_group_headers = self._get_col_group_headers(
            columns,
            len(measures),
            matrix['column_group_depth'],
            has_total_column,
        )

        return {
            'model': self.model_id,
            'title': self.title or self.favorite_filter_id.name or self.name,
            'col_group_headers': col_group_headers,
            'measure_headers': [
                {
                    'title': measure_labels[measure],
                    'width': 1,
                    'height': 1,
                    'is_bold': False,
                }
                for _key, _label in columns
                for measure in measures
            ] + ([{
                'title': measure_labels[measure],
                'width': 1,
                'height': 1,
                'is_bold': True,
            } for measure in measures] if has_total_column else []),
            'rows': [{
                'title': 'Total',
                'indent': 0,
                'values': [
                    {
                        'is_bold': True,
                        'value': column_totals[(col_key, measure)] if (col_key, measure) in column_totals else '',
                    }
                    for col_key, _label in columns
                    for measure in measures
                ] + ([{
                    'is_bold': True,
                    'value': grand_totals[measure] if measure in grand_totals else '',
                } for measure in measures] if has_total_column else []),
            }] + [
                {
                    'title': row_label,
                    'indent': row_indent,
                    'values': [
                        {
                            'is_bold': False,
                            'value': row_values[(row_key, col_key, measure)] if (
                                row_key, col_key, measure
                            ) in row_values else '',
                        }
                        for col_key, _label in columns
                        for measure in measures
                    ] + ([{
                        'is_bold': True,
                        'value': row_totals[(row_key, measure)] if (row_key, measure) in row_totals else '',
                    } for measure in measures] if has_total_column else []),
                }
                for row_key, row_label, row_indent in rows
            ],
            'measure_count': len(measures),
            'origin_count': 1,
        }

    def _get_col_group_headers(self, columns, measure_count, column_group_depth, has_total_column):
        if not columns:
            return [[]]
        # The first row is the root Total column group. Child rows are appended
        # below for each configured column groupby level.
        headers = [[
            {
                'title': 'Total',
                'width': len(columns) * measure_count,
                'height': 1,
                'is_bold': False,
            },
        ]]
        if has_total_column:
            headers[0].append({
                'title': '',
                'width': measure_count,
                'height': column_group_depth + 1,
                'is_bold': False,
            })

        for level in range(column_group_depth):
            header_row = []
            current_prefix = None
            current_title = ''
            current_width = 0
            for _key, label_path in columns:
                # Consecutive columns with the same prefix are collapsed into a
                # single header cell whose width spans all child measure cells.
                prefix = label_path[:level + 1]
                title = label_path[level]
                if current_prefix is None:
                    current_prefix = prefix
                    current_title = title
                    current_width = 1
                elif prefix == current_prefix:
                    current_width += 1
                else:
                    header_row.append({
                        'title': current_title,
                        'width': current_width * measure_count,
                        'height': 1,
                        'is_bold': False,
                    })
                    current_prefix = prefix
                    current_title = title
                    current_width = 1
            header_row.append({
                'title': current_title,
                'width': current_width * measure_count,
                'height': 1,
                'is_bold': False,
            })
            headers.append(header_row)
        return headers

    def _write_xlsx(self, matrix):  # noqa: C901
        # This follows Odoo's /web/pivot/export_xlsx writer shape closely so the
        # generated attachment looks like the standard pivot export.
        output = io.BytesIO()
        workbook = xlsxwriter.Workbook(output, {'in_memory': True})
        export_data = self._get_pivot_export_data(matrix)
        title = (export_data['title'] or self.name)[:31]
        worksheet = workbook.add_worksheet(title)

        header_bold = workbook.add_format({'bold': True, 'pattern': 1, 'bg_color': '#AAAAAA'})
        header_plain = workbook.add_format({'pattern': 1, 'bg_color': '#AAAAAA'})
        bold = workbook.add_format({'bold': True})
        rounded_number = workbook.add_format({'num_format': '0.##'})
        bold_rounded_number = workbook.add_format({'bold': True, 'num_format': '0.##'})

        measure_count = export_data['measure_count']
        origin_count = export_data['origin_count']
        col_group_headers = export_data['col_group_headers']

        x, y, carry = 1, 0, deque()
        for i, header_row in enumerate(col_group_headers):
            worksheet.write(i, 0, '', header_plain)
            for header in header_row:
                # carry stores headers with height > 1 so blank cells can be
                # written beneath them on following header rows.
                while carry and carry[0]['x'] == x:
                    cell = carry.popleft()
                    for j in range(measure_count * (2 * origin_count - 1)):
                        worksheet.write(y, x + j, '', header_plain)
                    if cell['height'] > 1:
                        carry.append({'x': x, 'height': cell['height'] - 1})
                    x = x + measure_count * (2 * origin_count - 1)
                for j in range(header['width']):
                    worksheet.write(y, x + j, header['title'] if j == 0 else '', header_plain)
                if header['height'] > 1:
                    carry.append({'x': x, 'height': header['height'] - 1})
                x = x + header['width']
            while carry and carry[0]['x'] == x:
                cell = carry.popleft()
                for j in range(measure_count * (2 * origin_count - 1)):
                    worksheet.write(y, x + j, '', header_plain)
                if cell['height'] > 1:
                    carry.append({'x': x, 'height': cell['height'] - 1})
                x = x + measure_count * (2 * origin_count - 1)
            x, y = 1, y + 1

        measure_headers = export_data['measure_headers']
        if measure_headers:
            worksheet.write(y, 0, '', header_plain)
            for measure in measure_headers:
                style = header_bold if measure['is_bold'] else header_plain
                worksheet.write(y, x, measure['title'], style)
                for i in range(1, 2 * origin_count - 1):
                    worksheet.write(y, x + i, '', header_plain)
                x = x + (2 * origin_count - 1)
            x, y = 1, y + 1
            worksheet.set_column(0, len(measure_headers), 16)

        for row in export_data['rows']:
            worksheet.write(y, 0, row['indent'] * '     ' + row['title'], header_plain)
            for cell in row['values']:
                value = cell['value']
                is_bold = cell.get('is_bold', False)
                fmt = bold if is_bold else None
                if isinstance(value, Number) and not isinstance(value, bool):
                    value = float(value)
                    if value.is_integer():
                        worksheet.write_number(y, x, int(value), fmt)
                    else:
                        worksheet.write_number(
                            y, x, round(value, 2),
                            bold_rounded_number if is_bold else rounded_number,
                        )
                else:
                    worksheet.write(y, x, value, fmt)
                x = x + 1
            x, y = 1, y + 1

        workbook.close()
        return output.getvalue()

    def _get_email_to(self):
        self.ensure_one()
        emails = [partner.email for partner in self.recipient_partner_ids if partner.email]
        emails.extend(email.strip() for email in (self.recipient_emails or '').split(',') if email.strip())
        # Preserve order while removing duplicates.
        return ','.join(dict.fromkeys(emails))

    def action_send_report(self):
        for report in self:
            email_to = report._get_email_to()
            if not email_to:
                continue
            matrix = report._read_pivot_matrix()
            xlsx_data = report._write_xlsx(matrix)
            name = report.title or report.favorite_filter_id.name or report.name
            filename = '%s.xlsx' % osutil.clean_filename('Pivot %s (%s)' % (name, report.model_id))
            attachment = self.env['ir.attachment'].sudo().create({
                'name': filename,
                'type': 'binary',
                'datas': base64.b64encode(xlsx_data).decode('ascii'),
                'mimetype': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                'res_model': report._name,
                'res_id': report.id,
            })
            if report.email_template_id:
                # Let the template render subject/body, but inject runtime
                # recipients and the freshly generated XLSX attachment.
                report.email_template_id.send_mail(
                    report.id,
                    force_send=True,
                    email_values={
                        'email_to': email_to,
                        'attachment_ids': [(4, attachment.id)],
                    },
                )
                continue
            mail = self.env['mail.mail'].sudo().create({
                'subject': name,
                'email_to': email_to,
                'body_html': '<p>Please find the sales report attached.</p>',
                'attachment_ids': [(4, attachment.id)],
            })
            mail.send()

    @api.model
    def _cron_send_daily_reports(self):
        reports = self.search([('active', '=', True), ('schedule', '=', 'daily')])
        reports.action_send_report()

    @api.model
    def _cron_send_weekly_reports(self):
        today_weekday = str(fields.Date.context_today(self).weekday())
        reports = self.search([
            ('active', '=', True),
            ('schedule', '=', 'weekly'),
            ('weekly_send_day', '=', today_weekday),
        ])
        reports.action_send_report()
