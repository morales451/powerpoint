import sys
from lxml import etree
A='{http://schemas.openxmlformats.org/drawingml/2006/main}'
# (prefix of original paragraph text, tightened text)
EDITS = {
 'slides/slide2.xml': [
  ('Consistent feedback is', "Consistent feedback: we’re difficult to deal with"),
  ('People held accountable', 'Hold people accountable'),
  ('Order taking is cumbersome', 'Cumbersome order taking'),
  ('CS sends an excel file', 'Customer asks for “24 drums of white silicone”; CS replies with a 200-SKU Excel file asking for the SAP #'),
  ('Take orders as they come', 'Take orders as written; call the customer if unclear. Hold customers to the Order Acknowledgement. Rely less on sales.'),
  ('Delayed responses', 'Slow responses'),
  ('Response times to inquiries', 'Inquiries (not orders) can take 24+ hours'),
  ('4-hour response time', '4-hour response on all emails'),
  ('Non-flexible to customers', 'Inflexible on requests (substitutions, alternate plants, expediting)'),
  ('If we’re out of PG988', 'If Kingman is out of PG988 but Carson has it, CS should get it to the customer fast, especially when a date is requested'),
  ('Give them access to inventory', 'Give CS inventory visibility across all plants. Train staff on approved substitutions.'),
  ('Some manufacturers sign-off', 'Some manufacturers sign off without walking the site'),
  ('Sign-off on warranties ASAP', 'Sign off ASAP; walk the roof within the roofer’s 2-year leak window. Batch sign-offs 3+ hours from base to cut travel.'),
  ('Building owners expect action', 'Building owners expect action within 24 hours'),
  ('Email acknowledgement in 4-hours', 'Email acknowledgement in 4 hours; site visit in 4 days'),
  ('Feedback is we are unreliable', 'Feedback: we’re unreliable'),
  ('Push roofers to distribution', 'Push roofers to distribution (solve pricing). Trackers on pallets. CS updates customers daily, unprompted.'),
  ('Material showing up', 'Material arrives days or weeks early with no one to accept it. Mainly hurts the roofer, but makes Henry/Carlisle look unprofessional too.'),
  ('Shipping at Henry’s convenience', 'Shipping at Henry’s convenience, not the customer’s'),
  ('If no delivery date is noted', 'With no requested date, we set one on the OA, then ship early without notice. Sometimes no one is there to receive it.'),
  ('CS should be communicating', 'Keep the customer informed. Ask before shipping early. Provide tracking and daily updates.'),
 ],
 'slides/slide3.xml': [
  ('Not having material ready', 'Material not ready to ship within 48 hours'),
  ('Not just a recent problem', 'Long-standing. Job-start items (primer, bleed blockers, mastics) must be in stock; trucks of 988 don’t help without primer.'),
  ('Higher minimums of accessories', 'Raise accessory minimums. Ramp production ahead of known storms and weather.'),
  ('Not being flexible on scope', 'Inflexible on scope/specs'),
  ('Many manufacturers require', 'Many manufacturers require “less.” We will flex.'),
  ('Failures will obviously', 'Failures kill future work with building owners'),
  ('Keep innovating', 'Keep innovating'),
  ('Missing Components to a System', 'Missing system components'),
  ('Missing a sprayable seam sealer', 'No sprayable seam sealer for silicone systems; no cementitious coating for Arizona'),
  ('Providing leads that don’t close', 'Leads that don’t close'),
  ('National account leads have', 'National account leads close well; Facebook/website leads often don’t'),
  ('We’re considered one of the better', 'We’re a strong lead provider, but Sherwin-Williams supplies more through paint contracts'),
  ('More collaboration with other divisions', 'Collaborate with other divisions. Free the national accounts team from firefighting and CS work to generate leads.'),
  ('Some manufacturers have a national agreement', 'Some manufacturers offer contractors discounted thermal surveys through national agreements'),
 ],
}
TREE = [
  ('How to Get and Build Upon', 'Winning and Growing Tier 1 Roofer Business'),
  ('Making customer provide SKU', 'Requiring SKU #s when customer already named the product'),
  ('Having to fill out their form', 'Requiring our form when customer already gave all the info'),
  ('Find ways to get product', 'Get product to customers so Henry never delays a job'),
  ('Does the warranty mean', 'Warranty matters to building owners'),
  ('Ease of getting warranty processed', 'Easy warranty processing'),
  ('Willing/ease of deviating', 'Flexibility on Henry standard spec'),
  ('Deliver material that is not damaged', 'Deliver material undamaged'),
]
EDITS['diagrams/data1.xml'] = TREE
EDITS['diagrams/drawing1.xml'] = TREE

root = sys.argv[1]
for part, edits in EDITS.items():
    path = f'{root}/ppt/{part}'
    tree = etree.parse(path)
    paras = [p for p in tree.iter(A+'p')]
    for prefix, new in edits:
        hits = [p for p in paras if ''.join(t.text or '' for t in p.iter(A+'t')).startswith(prefix)]
        assert hits, (part, prefix)
        for p in hits:
            runs = p.findall(A+'r')
            runs[0].find(A+'t').text = new
            for r in runs[1:]: p.remove(r)
    tree.write(path, xml_declaration=True, encoding='UTF-8', standalone=True)
    print(part, 'ok')
