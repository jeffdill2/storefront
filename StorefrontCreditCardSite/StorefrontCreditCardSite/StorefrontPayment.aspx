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
                left: 50px;
                top: 40px;
            }
        </style>
        
        <title>Storefront Credit Card Site</title>
    </head>
    
    <div style="width: 562px; height: 650px; padding: 10px; border: 5px solid gray; margin-left: 20px; ">
        <body style="background-color: #F8F8F8;">
            <div id="container">
                <img src="images/imghomeheader.jpg" id="img1" alt="Header Image" />
                <img src="images/HeeterWhiteLogo.png" id="img2" alt="Heeter Image" />
            </div>
        
            <br />
        
            <span style="font-family: Verdana, sans-serif; font-size: 13px; font-weight: normal;">
                <span style="margin-left: 10px; font-size: 24px; font-weight: bold; color: #585858;">Order Details</span><br />
            
                <br />
                <br />
            
                <table style="font-family: Verdana, sans-serif; font-size: 13px; margin-left: 30px; line-height: 20px; font-weight: bold; ">
                    <tr>
                        <td>Product Description :</td>
                        <td style="font-weight: normal; "><span runat="server" id="productDescriptionSpan" /></td>
                    </tr>

                    <tr>
                        <td>Print Quantity :</td>
                        <td style="font-weight: normal; "><span runat="server" id="printQuantitySpan" /></td>
                    </tr>

                    <tr>
                        <td>Per Piece Price :</td>
                        <td style="font-weight: normal; ">$<span runat="server" id="perPiecePriceSpan" /></td>
                    </tr>

                    <tr>
                        <td>&nbsp;</td>
                    </tr>

                    <!--
                    <tr>
                        <td>Order Fee Amount :</td>
                        <td>$<span runat="server" id="orderFeeSpan" /></td>
                    </tr>
                    -->

                    <tr>
                        <td>Print Charge Amount :</td>
                        <td style="font-weight: normal; ">$<span runat="server" id="printAmountSpan" /></td>
                    </tr>

                    <tr>
                        <td>Postage Charge Amount :</td>
                        <td style="font-weight: normal; ">$<span runat="server" id="postageAmountSpan" /></td>
                    </tr>

                    <tr>
                        <td>Shipping Charge Amount :</td>
                        <td style="font-weight: normal; ">$<span runat="server" id="shippingAmountSpan" /></td>
                    </tr>

                    <tr>
                        <td>Tax Charge Amount :</td>
                        <td style="font-weight: normal; ">$<span runat="server" id="taxAmountSpan" /></td>
                    </tr>

                    <tr>
                        <td>&nbsp;</td>
                    </tr>

                    <tr>
                        <td>&nbsp;</td>
                    </tr>

                    <tr>
                        <td style="line-height: 20px; font-size: 16px;">Total Charge Amount :&nbsp;&nbsp;</td>
                        <td style="font-weight: normal; line-height: 20px; font-size: 16px; ">$<span runat="server" id="totalChargeAmountSpan" /></td>
                    </tr>
                </table>

                <br />
                <br />
            
                <span style="margin-left: 10px;">
                    Please verify charge amounts, then click <i>Submit Payment</i><br />&nbsp;
                    to continue to a <b>secure</b> transaction page.
                </span>
            
                <br />
                <br />
                <br />
            </span>
        
            <!--
            FOR TESTING: https://test.authorize.net/gateway/transact.dll
            FOR LIVE: https://secure.authorize.net/gateway/transact.dll
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
                <input type='submit' runat="server" id='buttonLabel' style="text-align: center; width: 250px; height: 40px; margin-left: 150px; font-size: 14px; font-weight: bold; font-family: Verdana, sans-serif; border-style: outset;"/>
            </form>
        </body>
    </div>
</html>