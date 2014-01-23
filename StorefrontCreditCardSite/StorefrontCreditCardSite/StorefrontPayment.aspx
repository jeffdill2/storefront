<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="StorefrontPayment.aspx.cs" Inherits="StorefrontCreditCardSite.StorefrontPayment"%>

<%@ Import Namespace="System.Web"%>
<%@ Import Namespace="System.Web.UI"%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<meta http-equiv="cache-control" content="max-age=0" />
<meta http-equiv="cache-control" content="no-cache" />
<meta http-equiv="expires" content="0" />
<meta http-equiv="pragma" content="no-cache" />

<html xmlns="http://www.w3.org/1999/xhtml" >
    <head id="Head1" runat="server">
        <style type="text/css">
            body
            {
            }

            #container
            {
                position:relative;
            }

            #img2
            {
                position: absolute;
                left: 60px;
                top: 40px;
            }
        </style>
        
        <title>Storefront Credit Card Site</title>
    </head>
    
    <body style="background-color: #F8F8F8;">
        <div id="container">
            <img src="images/imghomeheader.jpg" id="img1" />
            <img src="images/HeeterWhiteLogo.png" id="img2" />
        </div>
        
        <br />
        
        <span style="font-family: Verdana; font-size: 13px; font-weight: normal;">
            <span style="margin-left: 10px; font-size: 24px; font-weight: bold; color: #585858;">Order Details</span><br />
            
            <br />
            <br />
            
            <span style="margin-left: 30px; line-height: 20px; font-weight: bold;">Product Description :&nbsp;&nbsp;&nbsp;</span><span runat="server" id="productDescriptionSpan"></span><br />
            <span style="margin-left: 30px; line-height: 20px; font-weight: bold;">Print Quantity :&nbsp;&nbsp;&nbsp;</span><span runat="server" id="printQuantitySpan"></span><br />
            <span style="margin-left: 30px; line-height: 20px; font-weight: bold;">Per Piece Price :&nbsp;&nbsp;&nbsp;</span>$<span runat="server" id="perPiecePriceSpan"></span><br />
            <br />
            <span style="margin-left: 30px; line-height: 20px; font-weight: bold;">Order Fee Amount :&nbsp;&nbsp;&nbsp;</span>$<span runat="server" id="orderFeeSpan"></span><br />
            <span style="margin-left: 30px; line-height: 20px; font-weight: bold;">Print Charge Amount :&nbsp;&nbsp;&nbsp;</span>$<span runat="server" id="printAmountSpan"></span><br />
            <span style="margin-left: 30px; line-height: 20px; font-weight: bold;">Postage Charge Amount :&nbsp;&nbsp;&nbsp;</span>$<span runat="server" id="postageAmountSpan"></span><br />
            
            <br />
            <br />
            
            <span style="margin-left: 10px; line-height: 20px; font-weight: bold; font-size: 16px;">Total Charge Amount :&nbsp;&nbsp;&nbsp;<span style="font-weight: normal;">$<span runat="server" id="totalChargeAmountSpan"></span></span></span><br />
            
            <br />
            <br />
            
            <span style="margin-left: 10px;">Please verify charge amounts, then click <i>Submit Payment</i></span><br />
            <span style="margin-left: 10px;">to continue to a <b>secure</b> transaction page.</span>
            
            <br />
            <br />
            <br />
        </span>
        
        <!--
        By default, this sample code is designed to post to our test server for
        developer accounts: https://test.authorize.net/gateway/transact.dll
        for real accounts (even in test mode), please make sure that you are
        posting to: https://secure.authorize.net/gateway/transact.dll
        -->
        
        <form id="simForm" runat="server" method='post' action='https://secure.authorize.net/gateway/transact.dll'>
            <input id="HiddenValue" type="hidden" value="Initial Value" runat="server" />
            <input type='hidden' runat="server" name='x_login' id='x_login' />
            <input type='hidden' runat="server" name='x_amount' id='x_amount' />
            <input type='hidden' runat="server" name='x_description' id='x_description' />
            <input type='hidden' runat="server" name='x_invoice_num' id='x_invoice_num' />
            <input type='hidden' runat="server" name='x_fp_sequence' id='x_fp_sequence' />
            <input type='hidden' runat="server" name='x_fp_timestamp' id='x_fp_timestamp' />
            <input type='hidden' runat="server" name='x_fp_hash' id='x_fp_hash' />
            <input type='hidden' runat="server" name='x_test_request' id='x_test_request' />
            <input type='hidden' name='x_show_form' value='PAYMENT_FORM' />
            <input type='hidden' runat="server" name='x_relay_response' id='x_relay_response' />
            <input type='hidden' runat="server" name='x_relay_url' id='x_relay_url' />
            <input type='hidden' runat="server" name='x_relay_always' id='x_relay_always' />
            <input type='submit' runat="server" id='buttonLabel' style="text-align: center; width: 200px; height: 30px; margin-left: 10px; font-family: Verdana; border-style: outset;"/>
        </form>
    </body>
</html>