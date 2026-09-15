const TOPICS = [
  "General inquiries",
  "Trade programme",
  "Payment received",
  "Expedition",
  "Payments",
  "Delivery confirmation",
  "Payment inquiries",
  "After sales",
  "December",
  "Other"
];

const REPLIES_VERSION = 2;

function r(id, topic, title, subject, body) {
  return { id, topic, title, subject, body: body.replace(/^\n+/, "").replace(/\n+$/, "") };
}

const DEFAULT_REPLIES = [
  r(
    "costing-in-progress",
    "General inquiries",
    "Costing in Progress",
    "Studio Delta — costing in progress",
    "Good day {{client_name}},\n\n" +
      "Thank you for your inquiry.\n\n" +
      "We’re happy to assist and have forwarded your specifications to our production team to generate a quote. I will send it through as soon as it’s ready.\n\n" +
      "To process your order smoothly, please provide your billing details (including VAT number, if applicable), delivery address, and contact number. Kindly also confirm the floor to which the order should be delivered; if this information is not provided, deliveries will automatically be made to the ground floor.\n\n" +
      "We also appreciate your interest in Studio Delta! We’d love to know how you first heard about us—Instagram, Google, or perhaps a recommendation from a friend. Your feedback is always valuable.\n\n" +
      "If you have any questions or need further assistance, please don’t hesitate to reach out.\n\n" +
      "Kind regards,"
  ),
  r(
    "information-required",
    "General inquiries",
    "Information Required",
    "Studio Delta — information required",
    "Good day {{client_name}},\n\n" +
      "Thank you for your inquiry.\n\n" +
      "Could you kindly confirm the following [details/specifications] or provide any additional information required?\n\n" +
      "To process your order smoothly, please provide your billing details (including VAT number, if applicable), delivery address, and contact number. Kindly also confirm the floor to which the order should be delivered; if this information is not provided, deliveries will automatically be made to the ground floor.\n\n" +
      "Once we have all the final specifications, we will generate a quote and send it through as soon as possible.\n\n" +
      "Thank you again for reaching out to Studio Delta. We’d love to know how you first heard about us—Instagram, Google, or a recommendation from a friend. Your feedback is always appreciated!\n\n" +
      "Looking forward to your response.\n\n" +
      "Kind regards,"
  ),
  r(
    "inquiry-with-colour",
    "General inquiries",
    "Inquiry with Colour",
    "Studio Delta — RAL colour",
    "Good day {{client_name}},\n\n" +
      "You can explore the RAL Classic colours at https://www.ralcolorchart.com/. Once you have selected a colour, please provide us with the corresponding RAL code to ensure accuracy. We also recommend cross-checking your chosen colour on multiple platforms to verify its appearance.\n\n" +
      "After we receive your confirmed RAL colour, we will check stock availability with our suppliers and suggest similar alternatives if your selected colour is unavailable.\n\n" +
      "Please note that achieving an exact match for custom colours can be challenging. It is the customer’s responsibility to review the selected RAL colour code across various platforms before providing final approval, as actual colours may vary from online images and representations. Studio Delta cannot be held liable if custom colours do not meet expectations.\n\n" +
      "If you have any questions or need further assistance, please don’t hesitate to reach out.\n\n" +
      "Kind regards,"
  ),
  r(
    "inquiry-with-wood",
    "General inquiries",
    "Inquiry with Wood",
    "Studio Delta — wood and stain options",
    "Good day {{client_name}},\n\n" +
      "We offer the following wood options for our units: American Ashwood, Rubberwood, Pine, and Saligna. Unfortunately, I do not have images of the raw wood available at the moment.When an order is placed, we procure the wood directly from our supplier and then stain it to the client’s preference.\n\n" +
      "You can view the various woods and stain options on the following website:\n" +
      "https://publicassets.rubiomonocoat.com/dam/spaces/a6adf4a8f8c149a59675432455a2506a\n" +
      "If you have trouble navigating the site or need assistance, feel free to confirm your availability, and I’d be happy to call you to guide you through the options.\n\n" +
      "A few notes on the wood options:\n\n" +
      "White Oak looks very similar to American Ash.\n" +
      "Hickory stains similarly to Rubberwood.\n" +
      "I wouldn’t recommend staining Saligna as its grain is quite straight, though the color would resemble that of Red Oak.\n" +
      "Please let me know if you have any questions or if you'd like further assistance.\n\n" +
      "Kind regards,"
  ),
  r(
    "standard-quotes",
    "General inquiries",
    "Standard Quotes",
    "Studio Delta — quotation {{quote_no}}",
    "Good day {{client_name}},\n\n" +
      "I hope this email finds you well.\n\n" +
      "Thank you for considering Studio Delta for your furniture needs. Please find attached the detailed quote for your inquiry. Before we proceed with manufacturing, we kindly ask that you review all the details carefully, including item descriptions, materials, and dimensions, to ensure everything aligns with your expectations.\n\n" +
      "Your satisfaction is our top priority. Confirming the accuracy of the information in the quote will allow our production team to create a product that meets your exact specifications.\n\n" +
      "When making your payment, please reference your quote number for smooth processing.\n\n" +
      "We truly appreciate your business and look forward to crafting a piece of furniture you’ll enjoy for years to come.\n\n" +
      "Kind regards,"
  ),
  r(
    "custom-quotes",
    "General inquiries",
    "Custom Quotes",
    "Studio Delta — custom quotation {{quote_no}}",
    "Good day {{client_name}},\n\n" +
      "I hope this email finds you well.\n\n" +
      "Thank you for considering Studio Delta for your furniture needs. Attached, you will find the detailed quote for your inquiry. Before we proceed with manufacturing, we kindly request that you review the information carefully—including item descriptions, materials, and dimensions—to ensure everything meets your expectations.\n\n" +
      "Your satisfaction is our top priority, and verifying the accuracy of the quote will help our production team deliver a product that aligns perfectly with your requirements.\n\n" +
      "To process your order smoothly, please provide your billing details (including VAT number, if applicable), delivery address, and contact number. Kindly also confirm the floor to which the order should be delivered; if this information is not provided, deliveries will automatically be made to the ground floor.\n\n" +
      "When making payment, please reference your quote number to ensure smooth processing.\n\n" +
      "We truly appreciate your business and look forward to creating a piece of furniture that you will cherish for years to come.\n\n" +
      "Kind regards,"
  ),
  r(
    "stock-quotes",
    "General inquiries",
    "Stock Quotes",
    "Studio Delta — stock quotation {{quote_no}}",
    "Good day {{client_name}},\n\n" +
      "I hope this email finds you well.\n\n" +
      "Thank you for considering Studio Delta for your furniture needs. Attached, please find the detailed quote for your inquiry. We kindly ask that you carefully review all details, including item descriptions, materials, and dimensions, to ensure everything aligns with your expectations.\n\n" +
      "Your satisfaction is our top priority, and confirming the accuracy of the quote will allow us to deliver a product that meets your exact specifications.\n\n" +
      "When making payment, please reference your quote number for smooth processing. Please note that payment must be completed within four hours, after which the unit will be offered to the next person in line.\n\n" +
      "Kind regards,"
  ),
  r(
    "quote-follow-up",
    "General inquiries",
    "Quote Follow Up",
    "Studio Delta — following up on {{quote_no}}",
    "Good day {{client_name}},\n\n" +
      "I hope this message finds you well.\n\n" +
      "I’m just sending a gentle reminder regarding the quote we previously sent. If you have any questions or require additional information, please don’t hesitate to let me know—I’d be happy to assist further.\n\n" +
      "I look forward to hearing from you.\n\n" +
      "Kind regards,"
  ),
  r(
    "trade-invitation",
    "Trade programme",
    "Trade programme invitation",
    "Invitation: Exclusive 10% Trade Discount – Let’s Get You Registered",
    "Good day {{client_name}},\n\n" +
      "I hope you are well.\n\n" +
      "We are excited to introduce the Studio Delta Trade Program, designed specifically for our interior design and architectural partners. The program offers:\n\n" +
      "An exclusive 10% trade discount on our handcrafted furniture\n\n" +
      "Priority quoting for your projects\n\n" +
      "Early previews of new collections\n\n" +
      "Access to finish samples and technical specifications\n\n" +
      "Customisation support for larger-scale projects\n\n" +
      "To make things easier, with your permission, we would be happy to begin the registration process on your behalf and set up your Trade Account so you can enjoy these benefits immediately.\n\n" +
      "Please let us know if you would like us to proceed with your registration, and we’ll take care of the rest.\n\n" +
      "We truly value our partnership and look forward to supporting your upcoming projects with refined, architectural furniture that elevates every space.\n\n" +
      "Warm regards,"
  ),
  r(
    "trade-application",
    "Trade programme",
    "Application in Progress",
    "Studio Delta — trade programme application",
    "Good day {{client_name}},\n\n" +
      "Thank you for submitting your trade program request we will be in touch shortly with the next steps once your account is set up.\n\n" +
      "Best regards,"
  ),
  r(
    "payment-gauteng",
    "Payment received",
    "Gauteng",
    "Studio Delta — payment received",
    "Good day {{client_name}},\n\n" +
      "Thank you for your purchase from Studio Delta.\n\n" +
      "We have received your payment, and production will begin as soon as possible. Our products are handcrafted in our workshop by our skilled team, and each piece requires a lead time to ensure the quality and craftsmanship we pride ourselves on. Your provisional delivery date is [insert date]. Our team will confirm the exact delivery date closer to the time.\n\n" +
      "Please note that, while we strive to deliver on time, we, along with our smaller suppliers, may experience occasional strains on our production line. We appreciate your patience as we work to get your order to you as soon as possible without compromising quality. Should there be any changes, we will communicate with you promptly.\n\n" +
      "We hope you’ll enjoy your unique, handcrafted furniture, proudly made in South Africa.\n\n" +
      "Thank you again for choosing Studio Delta and entrusting us with your order. We would love to know how you first heard about us—perhaps through Instagram, Google, or a recommendation from a friend? Your feedback is always greatly appreciated.\n\n" +
      "Wishing you a wonderful day ahead.\n\n" +
      "Kind regards,"
  ),
  r(
    "payment-out-gauteng",
    "Payment received",
    "Out of Gauteng",
    "Studio Delta — payment received",
    "Good day {{client_name}},\n\n" +
      "Thank you for your purchase from Studio Delta.\n\n" +
      "We have received your payment and will begin production as soon as possible. All our products are handcrafted in our workshop by our skilled team, and each item requires a lead time to ensure the care and craftsmanship that we are known for. Your provisional delivery date is during the week of [insert date]. Our team will be in touch closer to delivery.\n\n" +
      "Please be aware that, due to possible strains on our production line and with our smaller suppliers, there may be slight delays. We appreciate your patience as we work hard to get your items to you as quickly as possible, while maintaining our high standards of quality. Should anything change, we will promptly communicate with you.\n\n" +
      "We hope you’ll enjoy your uniquely handcrafted furniture piece, proudly made in South Africa.\n\n" +
      "Thank you again for choosing Studio Delta and entrusting us with your order. We’d love to hear how you first found us—maybe through Instagram, Google, or a recommendation from a friend? Your feedback is always greatly appreciated!\n\n" +
      "Wishing you a lovely day ahead.\n\n" +
      "Kind regards,"
  ),
  r(
    "expedition",
    "Expedition",
    "Expedition",
    "Studio Delta — expedition request",
    "Good day {{client_name}},\n\n" +
      "I hope you’re doing well.\n\n" +
      "Thank you for your email. While we understand the importance of having the unit delivered prior to XXX. I must inform you that we cannot promise an earlier delivery date.\n\n" +
      "All our products have specific lead times that consider previous orders, our production capacity, and timelines from our powder coaters and glass suppliers. We also include a buffer to address any potential issues, allowing sufficient time to resolve them.\n\n" +
      "Since this is a custom order, we are unable to use our standard templates, which means additional time is required to first build the frame and then place the order for the glass and mirror.\n\n" +
      "If we were to expedite your order, it could impact other scheduled orders, require overtime from our team, and incur additional fees for smaller suppliers and expedition services. Should you wish to proceed with an earlier delivery, please be aware that an expedition fee will apply.\n\n" +
      "I can only confirm delivery once the unit has been assembled and passed quality assurance. We recommend not scheduling any installation until the unit has been delivered to ensure everything is in order.\n\n" +
      "Thank you for your understanding, and please let me know how you would like to proceed.\n\n" +
      "Kind regards,"
  ),
  r(
    "request-pop",
    "Payments",
    "Request for Proof of Payment",
    "Studio Delta — proof of payment",
    "Good day {{client_name}},\n\n" +
      "Thank you for your order.\n\n" +
      "To help expedite the process, could you kindly send us your proof of payment? Once received, we will be able to begin production on your order.\n\n" +
      "Kind regards,"
  ),
  r(
    "request-float-link",
    "Payments",
    "Request for Float Link",
    "Studio Delta — Float payment link",
    "Good day {{client_name}},\n\n" +
      "Thank you for your feedback.\n\n" +
      "Please find the link to our Float payment page here: https://secure.float.co.za/float_anywhere/7dd7ee40-f79d-4c08-9413-deaffb0a5063/order_details/new.\n\n" +
      "When making the payment, kindly use the quote number as a reference to ensure proper processing.\n\n" +
      "If you need any further information or assistance, please don’t hesitate to reach out.\n\n" +
      "Kind regards,"
  ),
  r(
    "delivery-follow-up",
    "Delivery confirmation",
    "Follow Up On Provisional Delivery Date",
    "Studio Delta — provisional delivery date",
    "Good day {{client_name}},\n\n" +
      "I hope this message finds you well.\n\n" +
      "I wanted to update you regarding your order. I’m pleased to confirm that everything is still on track for delivery xxx.\n\n" +
      "Once your order has been fully assembled and passed through the necessary quality checks, I will reach out to confirm the exact delivery details.\n\n" +
      "Thank you for your patience and continued trust in our services. Should you have any questions in the meantime, please don't hesitate to get in touch.\n\n" +
      "Best regards,"
  ),
  r(
    "delivery-delays",
    "Delivery confirmation",
    "Delays",
    "Studio Delta — delivery date update",
    "Good day {{client_name}},\n\n" +
      "I hope you're doing well.\n\n" +
      "Due to increased pressure on our production line, including from some of our smaller suppliers, your provisional delivery date has been extended. We are now aiming to complete your order by [insert date]. If there are any further changes between now and your new provisional delivery date, I will inform you immediately.\n\n" +
      "Once your order has been assembled and passed our quality assurance checks, I will confirm the final delivery date with you.\n\n" +
      "We truly appreciate your patience and understanding during this time. Should you have any questions or need assistance, please don’t hesitate to reach out.\n\n" +
      "Wishing you a lovely day ahead.\n\n" +
      "Kind regards,"
  ),
  r(
    "delivery-gauteng",
    "Delivery confirmation",
    "Gauteng",
    "Studio Delta — delivery confirmation",
    "Good day {{client_name}},\n\n" +
      "We are pleased to confirm that your Studio Delta order will be delivered on the XXX.\n\n" +
      "Our team will reach out to you the day of delivery prior to departing our facility to confirm your availability and provide an estimated time of arrival, as they depart from our facility in Silverton\n\n" +
      "If your unit appears uneven or if the doors look misaligned, you can easily rectify this by turning the adjustable levelling feet or screws located at the base of the unit. Simply use your fingers to gently turn the levelling feet or screws clockwise or counterclockwise to adjust the height of each corner until the unit is level. Be sure to check the alignment of the doors after adjusting the feet. For a visual guide, please refer to our instructional video. https://www.youtube.com/watch?v=-ia0WGGFPPw.\n\n" +
      "Should you have any questions or need further assistance, feel free to reach out.\n\n" +
      "Wishing you a great day ahead.\n\n" +
      "Kind regards,"
  ),
  r(
    "delivery-out-gauteng",
    "Delivery confirmation",
    "Out of Gauteng",
    "Studio Delta — courier update",
    "Good day {{client_name}},\n\n" +
      "We’re pleased to inform you that your Studio Delta order is currently with our couriers.\n\n" +
      "The couriers will contact you 24 hours in advance to provide an estimated time of arrival (ETA).\n\n" +
      "If your unit appears uneven or if the doors look misaligned, you can easily rectify this by turning the adjustable levelling feet or screws located at the base of the unit. Simply use your fingers to gently turn the levelling feet or screws clockwise or counterclockwise to adjust the height of each corner until the unit is level. Be sure to check the alignment of the doors after adjusting the feet. For a visual guide, please refer to our instructional video. https://www.youtube.com/watch?v=-ia0WGGFPPw\n\n" +
      "If you have any questions, please don’t hesitate to reach out.\n\n" +
      "Wishing you a great day ahead.\n\n" +
      "Kind regards,"
  ),
  r(
    "deposit-query",
    "Payment inquiries",
    "Deposit Query",
    "Studio Delta — deposit",
    "Good day {{client_name}},\n\n" +
      "Thank you for reaching out regarding the deposit for one of our units.\n\n" +
      "As per our standard practice, we generally require full payment upfront for orders below R75,000. However, we do offer Mobicred as a flexible payment option on our website, allowing you to spread the total cost over a chosen period.\n\n" +
      "If this option does not suit your needs, please let me know, and I can escalate your request to management to explore any alternative arrangements that might work better for you.\n\n" +
      "Feel free to reach out if you need any further assistance or have additional questions.\n\n" +
      "Kind regards,"
  ),
  r(
    "google-review",
    "After sales",
    "Google Review (positive clients only)",
    "Studio Delta — we would love your review",
    "Good day {{client_name}},\n\n" +
      "I hope this email finds you well.\n\n" +
      "Thank you again for your purchase from Studio Delta. We always love seeing our customers' pieces in their homes and hearing any feedback you may have.\n\n" +
      "If you have any photos or videos of your new furniture that you’re willing to share, we would greatly appreciate it! Alternatively, if you’d like to leave a quick online review, it would mean a lot to us.\n\n" +
      "Using your Gmail account, it should only take a moment—simply click here to leave your review:\n" +
      "https://g.page/r/CezbgV6pm08TEAI/review\n\n" +
      "Thank you in advance, and we look forward to seeing you again soon!\n\n" +
      "Kind regards,"
  ),
  r(
    "december-payment",
    "December",
    "Payment Received",
    "Studio Delta — payment received",
    "Good day {{client_name}},\n\n" +
      "Thank you for your recent purchase from Studio Delta. We have noted your specifications and billing details, and I am pleased to confirm that your payment has been successfully received. Your order will soon enter production, and our team is committed to crafting each piece with the quality and attention to detail that we’re known for.\n\n" +
      "As we approach the holiday season, please be aware that many of our suppliers will be closing for the holidays starting December 12. While our standard lead time is xxx, year-end closures may result in delays for new orders. We will have a dedicated team working during the holiday period to fulfil as many orders as possible. Your provisional delivery date is currently set for xxx. Should there be any changes, we will keep you informed and updated.\n\n" +
      "We truly appreciate your patience and understanding, and we are grateful to have you as a customer. Your unique, handcrafted piece will certainly be worth the wait, and we look forward to delivering it to you.\n\n" +
      "Additionally, if you could kindly let us know how you first heard about Studio Delta, we would greatly appreciate your feedback as it helps us improve our services.\n\n" +
      "Thank you once again for choosing Studio Delta.\n\n" +
      "Kind regards,"
  ),
  r(
    "december-delays",
    "December",
    "Delays",
    "Studio Delta — December delay",
    "Good day {{client_name}},\n\n" +
      "I trust this message finds you well. I am writing to you regarding your recent order, and I deeply regret to inform you that we will unfortunately be unable to fulfil the delivery within this year as initially planned.\n\n" +
      "We are currently experiencing delays with both our team and several smaller suppliers. Despite our best efforts to meet the original deadline, these delays, combined with the closure of our suppliers from December 13th until the new year, have made it impossible to complete the order in time. Additionally, our couriers are also closed during this period, which further contributes to the challenge.\n\n" +
      "While we will have our team working throughout the festive season, we are still dependent on key processes such as powder coating the frame and receiving the glass delivery, which are affected by the supplier backlogs.\n\n" +
      "As a result, we now anticipate that your order will be delivered during the week of January 16th. We fully understand the inconvenience this causes and sincerely apologize for the entire month-long delay. Please be assured that we are doing everything we can to expedite the process and ensure all orders are fulfilled as quickly as possible.\n\n" +
      "We are truly sorry for any disruption this may cause and greatly appreciate your understanding during this time.\n\n" +
      "Kind regards,"
  ),
  r(
    "december-master-movers",
    "December",
    "Master Movers",
    "Studio Delta — Master Movers delivery",
    "Good day {{client_name}},\n\n" +
      "I trust this email finds you well.\n\n" +
      "Regarding your order, we have arranged for it to be delivered via Master Movers. They will contact you directly 24 hours prior to delivery to provide an estimated time of arrival (ETA) and confirm your availability to receive the delivery.\n\n" +
      "If you wish to follow up with them, you can contact their office at (021) 534 1582. Please provide them with your name, address, and let them know you are expecting a delivery from Studio Delta. This will allow them to give you an update on your delivery status.\n\n" +
      "Wishing you a wonderful holiday season and a happy New Year! We greatly appreciate your continued support and look forward to working with you in the year ahead.\n\n" +
      "Kind regards,"
  ),
  r(
    "december-campos",
    "December",
    "Campos",
    "Studio Delta — Campos delivery",
    "Good day {{client_name}},\n\n" +
      "I trust this email finds you well.\n\n" +
      "Regarding your order, we have arranged for it to be delivered via Campos. They will contact you directly 24 hours prior to delivery to provide an estimated time of arrival (ETA) and confirm your availability to receive the delivery.\n\n" +
      "If you would like to follow up with them, you may contact their office at (011) 608 6400. Please provide them with your name, address, and inform them that you are expecting a delivery from Studio Delta. This information will allow them to give you an update on your delivery status.\n\n" +
      "Wishing you a wonderful holiday season and a happy New Year! We greatly appreciate your continued support and look forward to working with you in the year ahead.\n\n" +
      "Kind regards,"
  ),
  r(
    "december-mds",
    "December",
    "MDS",
    "Studio Delta — waybill {{quote_no}}",
    "Good day {{client_name}},\n\n" +
      "Your waybill number is [XXX], and you may track your order using the following website: https://collivery.net/tracking.\n\n" +
      "We truly appreciate your patience and would love to receive images of the unit in your space once it has been installed.\n\n" +
      "Wishing you a wonderful holiday season and a happy New Year! We greatly appreciate your continued support and look forward to working with you in the year ahead.\n\n" +
      "Kind regards,"
  ),
  r(
    "december-payment-followup",
    "December",
    "Payment Followup",
    "Studio Delta — following up on {{quote_no}}",
    "Good day {{client_name}},\n\n" +
      "I hope you are well. I am reaching out to follow up on the quote we recently sent to you.\n\n" +
      "Is there any additional information or clarification we can provide you with at this time?\n\n" +
      "Unfortunately, all orders placed from today onward will only be delivered in the New Year. We estimate delivery to take place between the last week of January and the first week of February.\n\n" +
      "While many of our smaller suppliers will be closed over the festive season, we will have a dedicated team working throughout this period to ensure there are no delays once our suppliers reopen.\n\n" +
      "Final delivery dates will be communicated during the second week of January.\n\n" +
      "We appreciate your understanding during this time and look forward to any further questions or requests you may have.\n\n" +
      "Kind regards,"
  ),
  r(
    "lead-times",
    "Other",
    "Lead Times",
    "Studio Delta — lead times",
    "Good day {{client_name}},\n\n" +
      "I hope you are doing well.\n\n" +
      "We want to assure you that each piece of furniture we create is handcrafted with care, which naturally extends the production time. We also work with smaller suppliers whose timelines can affect our overall lead times.\n\n" +
      "We do have this information available on our website, but we understand that it might not be as visible as it should be. We will communicate with our web team to enhance its visibility for future customers.\n\n" +
      "You can find our current lead times detailed at https://www.studiodelta.co.za/faq/. If you have any further questions or need assistance, please feel free to reach out.\n\n" +
      "Kind regards,"
  ),
  r(
    "abandoned-cart",
    "Other",
    "Abandoned Cart",
    "Studio Delta — your checkout",
    "Good day {{client_name}},\n\n" +
      "We hope you’re having a great day! 🙂\n\n" +
      "It looks like you didn’t have a chance to complete your checkout. If you have any questions about our products, shipping, or delivery, please don’t hesitate to reply to this email. We’re here to assist you every step of the way!\n\n" +
      "We look forward to hearing from you soon.\n\n" +
      "Kind regards,"
  ),
  r(
    "referral-letters",
    "Other",
    "Referral Letters",
    "Studio Delta — reference letter",
    "Dear {{client_name}},\n\n" +
      "I hope this email finds you well. \n\n" +
      "As a valued client of Studio Delta, we greatly appreciate the trust you placed in us for your furniture needs. It has been a pleasure working with you, and we hope the furniture we provided has met or exceeded your expectations. \n\n" +
      "We are currently reaching out to some of our valued clients to request reference letters that speak to their experiences with our products and services. We believe your insights could provide valuable perspectives for potential clients.\n\n" +
      "If you could spare a few moments to share your thoughts on your experience, we would be extremely grateful.  We have gone through the liberty of setting up a draft however, you are welcome to edit it if you wish and if you could please add your letterhead and sign that would be great.\n\n" +
      "We understand that your time is precious, and we genuinely appreciate your consideration of this request. If you have any specific requirements or guidelines for the reference letter, please let us know, and we will ensure they are met. \n\n" +
      "Should you agree to provide a reference letter, please feel free to send it directly to info@studiodelta.co.za. We plan to use these reference letters for promotional purposes, and your consent to do so would be greatly appreciated. \n\n" +
      "Thank you once again for choosing Studio Delta. We look forward to hearing from you and sincerely appreciate your ongoing support.\n\n" +
      "Kind regards,"
  )
];

module.exports = { TOPICS, REPLIES_VERSION, DEFAULT_REPLIES };
