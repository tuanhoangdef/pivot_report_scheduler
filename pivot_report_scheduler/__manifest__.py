{
    'name': 'Pivot Report Scheduler',
    'version': '19.0.1.0.0',
    'category': 'Reporting',
    'license': 'LGPL-3',
    'summary': 'Schedule Odoo pivot favorites and email XLSX reports automatically.',
    'description': 'Schedule saved Odoo pivot favorites and send them by email as XLSX reports.',
    'author': 'tuanhoangdef <hng.atuan@gmail.com>',
    'price': 0,
    'currency': 'USD',
    'depends': [
        'base',
        'mail',
        'sale',
    ],
    'data': [
        'security/ir.model.access.csv',
        'data/scheduled_pivot_report_data.xml',
        'views/scheduled_pivot_report_views.xml',
    ],
    'images': [
        'static/description/icon.png',
    ],
    'website': 'https://github.com/tuanhoangdef/pivot_report_scheduler/tree/19.0/',
    'installable': True,
    'auto_install': False
}
