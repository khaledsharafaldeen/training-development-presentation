"""generate_pptx.py

Script to generate a ready .pptx file for "استراتيجية التدريب والتطوير في الإدارة العامة".

Usage:
1. Create a Python virtualenv and install dependencies:
   python -m venv venv
   source venv/bin/activate  # on Windows: venv\Scripts\activate
   pip install python-pptx

2. Run the script:
   python generate_pptx.py

Output:
- presentation/output/Training_Development_Strategy_Khaled_Sharaf_Aldeen.pptx

Notes:
- The script sets the Arabic font to 'Noto Sans Arabic'. Ensure the font is installed on your system so PowerPoint displays text correctly.
- If you want a different font, edit FONT_NAME variable.

This script programmatically creates 16 slides with titles, bullet points and speaker notes based on the repository's markdown content.
"""
from pptx import Presentation
from pptx.util import Pt, Inches
from pptx.enum.text import PP_ALIGN
import os

FONT_NAME = 'Noto Sans Arabic'  # change if you prefer another installed Arabic font
OUTPUT_DIR = 'presentation/output'
OUTPUT_FILE = os.path.join(OUTPUT_DIR, 'Training_Development_Strategy_Khaled_Sharaf_Aldeen.pptx')

slides_content = [
    {
        'title': 'استراتيجية التدريب والتطوير في الإدارة العامة',
        'bullets': ['بناء الكفاءات وتحسين الأداء المؤسسي'],
        'notes': 'ترحيب قصير وذكر هدف العرض. مقدم العرض: خالد شرف الدين — الجهة: وزارة'
    },
    {
        'title': 'مقدمة وسياق',
        'bullets': [
            'التدريب والتطوير ركيزة لتحسين كفاءة الإدارة العامة.',
            'التحولات التقنية والتنظيمية تتطلب تطوير مهارات الموظفين.',
            'الاستثمار في الموارد البشرية يؤثر إيجابياً على جودة الخدمات.'
        ],
        'notes': 'اذكر انهيار سريع للواقع الحالي مع أمثلة من الإدارة المالية العامة.'
    },
    {
        'title': 'أهداف الاستراتيجية',
        'bullets': [
            'رفع كفاءة الأداء الإداري والقيادي.',
            'بناء قدرات التخطيط وصنع القرار.',
            'تعزيز ثقافة الخدمة العامة والشفافية.',
            'دعم التحول الرقمي والابتكار المؤسسي.'
        ],
        'notes': 'اربط كل هدف بمؤشر قياس قابل للقياس.'
    },
    {
        'title': 'تحليل الوضع الحالي (SWOT مختصر)',
        'bullets': [
            'نقاط القوة: خبرة مهنية لدى كوادر مالية وإدارية، هياكل تشغيلية واضحة.',
            'نقاط الضعف: فجوات مهارية، محدودية أدوات القياس.',
            'الفرص: شراكات تدريبية ومبادرات رقمنة.',
            'التهديدات: دوران موظفين وقيود الميزانية.'
        ],
        'notes': 'استخدم أمثلة عملية من واقع الإدارة المالية العامة.'
    },
    {
        'title': 'إطار الكفاءات والفئات المستهدفة',
        'bullets': [
            'الفئات: موظفو الصف الأمامي، الفنيون والماليون، المدراء الوسطيون، القيادات العليا.',
            'الكفاءات: فنية، قيادية، وسلوكيات مؤسسية.'
        ],
        'notes': 'ربط الإطار بالمسميات الوظيفية في الإدارة المالية.'
    },
    {
        'title': 'محاور الاستراتيجية',
        'bullets': [
            'المحور الأكاديمي: تحليل الاحتياجات، تصميم مناهج مبنية على نظريات.',
            'المحور الاحترافي: برامج تطبيقية، تدريب عملي، شراكات.'
        ],
        'notes': 'التوليف بين الأساس النظري والتطبيق المهني.'
    },
    {
        'title': 'تصميم البرامج والمحتوى',
        'bullets': [
            'أهداف تعليمية واضحة لكل برنامج.',
            'وحدات نظرية وتطبيقية متوازنة.',
            'حالات دراسية من الإدارة المالية العامة.'
        ],
        'notes': 'مثال: برنامج قيادي لمدة 6 أسابيع.'
    },
    {
        'title': 'أساليب التعلم والتقنيات',
        'bullets': [
            'التعلم الوجاهي: ورش ومحاكاة.',
            'التعلم الإلكتروني: microlearning ومنصة LMS.',
            'التعلم العملي: مشاريع داخل القسم المالي.',
            'التوجيه Coaching/Mentoring.'
        ],
        'notes': 'كيف تدمج الأساليب لتناسب موظفي الإدارة المالية.'
    },
    {
        'title': 'منهجية القياس (أكاديمي ومهني)',
        'bullets': [
            'قياس المعرفة: اختبارات قبل/بعد.',
            'قياس الكفاءات التطبيقية: ملاحظات المشرفين، عينات العمل.',
            'مؤشرات أداء مؤسسية: دقة التقارير، زمن المعالجة.'
        ],
        'notes': 'صل المؤشرات بمخرجات الإدارة المالية مثل جودة إعداد الميزانيات.'
    },
    {
        'title': 'نموذج كيركباتريك (مفصل)',
        'bullets': [
            'المستوى 1 (التفاعل): استبيان رضا — هدف: 85% بدرجة 4+.',
            'المستوى 2 (التعلم): اختبار قبل/بعد — هدف: زيادة 30%.',
            'المستوى 3 (السلوك): تقييم بعد 3 أشهر — هدف: 50% تطبيق مهارة.',
            'المستوى 4 (النتائج): مؤشرات مالية — هدف: خفض زمن المعالجة 15%.'
        ],
        'notes': 'ربط نتائج التدريب بالمؤشرات التنظيمية.'
    },
    {
        'title': 'خطة تنفيذ (12 شهرًا)',
        'bullets': [
            'التحضير: تحليل احتياجات (شهران).',
            'التصميم: إعداد المناهج (شهران).',
            'تجربة تجريبية: شهر.',
            'تنفيذ واسع: 6 أشهر.',
            'المتابعة والتقييم: طوال العام.'
        ],
        'notes': 'اعرض مخطط Gantt مبسط.'
    },
    {
        'title': 'الموارد والميزانية التقديرية',
        'bullets': [
            'تطوير المحتوى ومنصة LMS.',
            'أتعاب المدربين والاستشاريين.',
            'لوجستيات: قاعات ومواد.',
            'أدوات التقييم والمتابعة.'
        ],
        'notes': 'أضف جدول خيارين: أساسي وموسع.'
    },
    {
        'title': 'إدارة التغيير والتواصل',
        'bullets': [
            'خطة تواصل داخلية.',
            'إشراك القيادات.',
            'حوافز للمتميزين.'
        ],
        'notes': 'استخدم رسائل نموذجية للقادة.'
    },
    {
        'title': 'المخاطر وخطة التخفيف',
        'bullets': [
            'مقاومة التغيير.',
            'نقص التمويل.',
            'ضعف المتابعة.'
        ],
        'notes': 'إجراءات التخفيف: حملات توعية، خيارات مرنة، تقارير ربع سنوية.'
    },
    {
        'title': 'مؤشرات الأداء الرئيسية (KPIs)',
        'bullets': [
            'نسبة إكمال الدورات.',
            'نسبة رضا المتدربين.',
            'معدل تطبيق المهارات.',
            'تحسن زمن إنجاز العمليات.'
        ],
        'notes': 'حدد أهداف رقمية لكل KPI لمرحل 3/6/12 شهراً.'
    },
    {
        'title': 'الخلاصة والدعوة لاتخاذ قرار',
        'bullets': [
            'ملخص الأهداف والمنهجية.',
            'القرارات المطلوبة: اعتماد ميزانية تجريبية، تشكيل فريق تنفيذ.',
            'خطوات 30/60/90 يوماً.'
        ],
        'notes': 'اطلب من الحضور الموافقة أو تحديد مشاورات.'
    }
]


def create_presentation():
    prs = Presentation()
    # set slide dimensions if needed

    for item in slides_content:
        # Use title and content layout (layout 1 usually Title and Content)
        slide_layout = prs.slide_layouts[1]
        slide = prs.slides.add_slide(slide_layout)
        title = slide.shapes.title
        body = slide.shapes.placeholders[1]

        title.text = item['title']

        tf = body.text_frame
        tf.clear()
        for i, bullet in enumerate(item['bullets']):
            if i == 0:
                p = tf.paragraphs[0]
                p.text = bullet
            else:
                p = tf.add_paragraph()
                p.text = bullet
            p.level = 0
            for run in p.runs:
                run.font.name = FONT_NAME
                run.font.size = Pt(20)
                run.font.bold = False

        # set title font
        for paragraph in title.text_frame.paragraphs:
            for run in paragraph.runs:
                run.font.name = FONT_NAME
                run.font.size = Pt(28)
                run.font.bold = True

        # Add speaker notes
        notes_slide = slide.notes_slide
        notes_text_frame = notes_slide.notes_text_frame
        notes_text_frame.text = item.get('notes', '')

    # Remove default first blank slide if exists (python-pptx starts with a blank slide)
    if len(prs.slides) > len(slides_content):
        # Keep only the added slides by creating a new Presentation and copying slides is complex;
        # Instead, create presentation from scratch without default slide by saving a new one created above.
        pass

    os.makedirs(OUTPUT_DIR, exist_ok=True)
    prs.save(OUTPUT_FILE)
    print(f'Presentation saved to: {OUTPUT_FILE}')


if __name__ == '__main__':
    create_presentation()
